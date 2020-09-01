using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using App = MyJournal.Notebook.API.ApplicationExtensions;

// Internals of this component are not visible to COM
[assembly: ComVisible(false)]

// Guid and Version for the type library (.tlb)
[assembly: Guid("2F2C9D5C-3B72-4592-ABAC-F354C7B9F18B")]
[assembly: TypeLibVersion(1, 0)]

namespace MyJournal.Notebook
{
    static class Component
    {
#if WIN32
        const string
          ONENOTE_APP_NOT_FOUND = "OneNote 32-bit application not found.";
#elif WIN64
        const string
          ONENOTE_APP_NOT_FOUND = "OneNote 64-bit application not found.";
#endif
        internal const string Description = "My Journal Notebook COM Add-in";
        internal const string FriendlyName = "My Journal";
        internal const string ProgId = "MyJournal.Notebook.Connect";

        // GUID for MyJournal.Notebook.Connect class
        internal const string ProgId_Guid = "B899BB4F-3A1E-4E6E-9040-9B9B65969180";

        internal const int CommandLineSafe = 0;
        internal const int LoadBehavior = 3;

        static readonly string s_appId_guid;
        static readonly string s_addIn_subkey;
        static readonly string s_appId_subkey;
        static readonly string s_clsId_subkey;

        #region Properties

        internal static string AppDataPath { get { return GetAppDataPath(); } }

        internal static readonly string AppIcon_ResourceName;

        internal static readonly AppSettings AppSettings;

        internal static readonly AssemblyInfo AssemblyInfo;

        internal static readonly string UserConfigPath;

        #endregion

        static Component()
        {
            s_appId_guid = string.Format("{{{0}}}", Component.ProgId_Guid);

            s_addIn_subkey =
                @"SOFTWARE\Microsoft\Office\OneNote\AddIns\" + Component.ProgId;

            s_appId_subkey = @"AppID\" + s_appId_guid;
            s_clsId_subkey = @"CLSID\" + s_appId_guid;

            AppIcon_ResourceName =
                string.Concat(typeof(Component).Namespace, ".App.ico");

            AppSettings = new AppSettings();
            AssemblyInfo = new AssemblyInfo();

            UserConfigPath = ConfigurationManager.OpenExeConfiguration(
              ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;
        }

        static string GetAppDataPath()
        {
            if (s_appDataPath == null)
            {
                var localAppData = Environment.SpecialFolder.LocalApplicationData;
                var folder = Environment.GetFolderPath(localAppData);
                lock (s_syncAppDataPath)
                {
                    if (s_appDataPath == null)
                    {
                        s_appDataPath =
                          Path.Combine(folder, AssemblyInfo.CompanyName, FriendlyName);
                    }
                }
                if (!Directory.Exists(s_appDataPath))
                {
                    Directory.CreateDirectory(s_appDataPath);
                }
            }
            return s_appDataPath;
        }

        internal static void Initialize()
        {
            Tracer.Initialize();
        }

        static bool IsOneNoteInstalled()
        {
            var isInstalled = false;
            var line = ONENOTE_APP_NOT_FOUND;

            if (App.IsOneNoteInstalled())
            {
                var exeFilePath = App.GetExeFilePath();
                var exeFileInfo = new Utils.ExeFileInfo(exeFilePath);
#if DEBUG
                var nl = new[] { Environment.NewLine };
                var options = StringSplitOptions.None;
                var args = exeFileInfo.ToString().Split(nl, options);
                foreach (var arg in args) Tracer.WriteDebugLine(arg);
#endif
                line = string.Format("{0} ({1}) is installed",
                    App.FriendlyName, exeFileInfo.ImageType);
#if WIN32
                isInstalled = exeFileInfo.Is32Bit();
#elif WIN64
                isInstalled = exeFileInfo.Is64Bit();
#endif
            }
            Tracer.WriteTraceMethodLine(line);
            return isInstalled;
        }

        internal static void Register(Type t)
        {
            Tracer.WriteTraceMethodLine();
            if (t.FullName == Component.ProgId)
            {
                if (!IsOneNoteInstalled())
                {
                    var errorMessage =
                      "Oops... Can't register component.\r\n" + ONENOTE_APP_NOT_FOUND;
                    Utils.WinHelper.DisplayError(errorMessage);
                    return;
                }
                WriteAssemblyInfoProperties();
                try
                {
                    Tracer.WriteDebugLine("Creating HKCR subkey: {0}", s_appId_subkey);
                    using (var k = Registry.ClassesRoot.CreateSubKey(s_appId_subkey))
                    {
                        k.SetValue(null, AssemblyInfo.Description);

                        // Use the default COM Surrogate process (dllhost.exe) to activate the DLL
                        k.SetValue("DllSurrogate", string.Empty);
                    }
                    Tracer.WriteDebugLine("Updating HKCR subkey: {0}", s_clsId_subkey);
                    using (var k = Registry.ClassesRoot.OpenSubKey(s_clsId_subkey, true))
                    {
                        k.SetValue("AppID", s_appId_guid);
                        using (var defaultIcon = k.CreateSubKey("DefaultIcon"))
                        {
                            var path = Assembly.GetExecutingAssembly().Location;
                            var resourceID = 0;
                            defaultIcon.SetValue(null, $"\"{path}\",{resourceID}");
                        }
                    }
                    // Register add-in for all users
                    Tracer.WriteDebugLine("Creating HKLM subkey: {0}", s_addIn_subkey);
                    using (var k = Registry.LocalMachine.CreateSubKey(s_addIn_subkey))
                    {
                        var dword = RegistryValueKind.DWord;
                        k.SetValue("CommandLineSafe", Component.CommandLineSafe, dword);
                        k.SetValue("Description", Component.Description);
                        k.SetValue("FriendlyName", Component.FriendlyName);
                        k.SetValue("LoadBehavior", Component.LoadBehavior, dword);
                    }
                }
                catch (Exception e)
                {
                    Utils.WinHelper.DisplayError("Oops... Can't register component.");
                    Utils.ExceptionHandler.HandleException(e);
                }
            }
        }

        internal static void Unregister(Type t)
        {
            Tracer.WriteTraceMethodLine();
            if (t.FullName == Component.ProgId)
            {
                WriteAssemblyInfoProperties();
                try
                {
                    // Remove any USER settings first, then HKLM and HKCR
                    var localUsers = Utils.WMI.GetLocalUserAccounts();
                    foreach (ManagementObject user in localUsers)
                    {
                        var sid = user["SID"].ToString();
                        using (var k = Registry.Users.OpenSubKey(sid))
                        {
                            if (k == null) continue; // user not logged on
                        }
                        var userSubKey = string.Format(@"{0}\{1}", sid, s_addIn_subkey);
                        using (var k = Registry.Users.OpenSubKey(userSubKey))
                        {
                            if (k == null) continue; // subkey is not found
                        }
                        Tracer.WriteDebugLine("Deleting USER {0} subkey: {1}", user["Name"], userSubKey);
                        Registry.Users.DeleteSubKey(userSubKey);
                    }
                    Tracer.WriteDebugLine("Deleting HKLM subkey: {0}", s_addIn_subkey);
                    Registry.LocalMachine.DeleteSubKey(s_addIn_subkey, false);

                    string[] collection = { s_appId_subkey, s_clsId_subkey };
                    foreach (var item in collection)
                    {
                        Tracer.WriteDebugLine("Deleting HKCR subkey: {0}", item);
                        Registry.ClassesRoot.DeleteSubKeyTree(item, false);
                    }
                }
                catch (Exception e)
                {
                    Utils.WinHelper.DisplayError("Oops... Can't unregister component.");
                    Utils.ExceptionHandler.HandleException(e);
                }
            }
        }

        /// <summary>
        /// Writes the global assembly attributes to the trace listeners.
        /// </summary>
        static void WriteAssemblyInfoProperties()
        {
            const int Count = 14;
            var spacer = new string('=', Count);
            Tracer.WriteDebugLine($"{spacer}[ Assembly Info ]{spacer}");

            var collection = AssemblyInfo.AssemblyAttributes
              .Where(x => !string.IsNullOrEmpty(x.Value));

            foreach (var kvp in collection)
            {
                Tracer.WriteDebugLine($"{kvp.Key,-Count}: {kvp.Value}");
            }
        }

        static volatile string s_appDataPath;
        static readonly object s_syncAppDataPath = new object();
    }
}
