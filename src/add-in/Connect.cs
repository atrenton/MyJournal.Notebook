using System;
using System.Runtime.InteropServices;
using Extensibility;
using MyJournal.Notebook.Diagnostics;
using Office = Microsoft.Office.Core;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook
{
    [ComVisible(true), Guid(ComClassId.Connect), ProgId(Component.ProgId)]
    public sealed class Connect : API.AddInBase, Extensibility.IDTExtensibility2,
      Office.IRibbonExtensibility

    {
        /// <summary>
        /// Implements the IDTExtensibility2 interface for this OneNote add-in
        /// component.
        /// <para>
        /// SEE: https://learn.microsoft.com/en-us/dotnet/api/extensibility.idtextensibility2?view=visualstudiosdk-2022
        /// </para>
        /// </summary>
        public Connect() : base()
        {
            var process = System.Diagnostics.Process.GetCurrentProcess();
            Tracer.WriteTraceTypeLine("DllHost PID = " + process.Id);
        }

        static Connect()
        {
            Component.Initialize();
        }

        #region COM Add-in Registration Methods

        [ComRegisterFunction()]
        public static void RegisterComponent(Type t)
        {
            ArgumentNullException.ThrowIfNull(t);

            var p = System.Diagnostics.Process.GetCurrentProcess();

            // Do not register if invoked by WiX Toolset Harvester (heat.exe)
            if (p.ProcessName != "heat") Component.Register(t);
        }

        [ComUnregisterFunction()]
        public static void UnregisterComponent(Type t)
        {
            ArgumentNullException.ThrowIfNull(t);

#if NETCOREAPP
            System.Runtime.Loader.AssemblyLoadContext.Default.Resolving +=
                Component.LoadFromExecutingAssemblyLocation;
#endif

            Component.Unregister(t);
        }

        #endregion

        #region IDTExtensibility2 Members
        // SEE: https://learn.microsoft.com/en-US/previous-versions/office/troubleshoot/office-developer/office-com-add-in-using-visual-c#2

        /// <summary>
        /// Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        /// Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        /// Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
            Tracer.WriteTraceMethodLine();

            // If add-in was connected with the COM Add-ins dialog, need to manually
            // fire OnStartupComplete event
            if (ConnectMode != ext_ConnectMode.ext_cm_Startup)
            {
                ((IDTExtensibility2)this).OnStartupComplete(ref custom);
            }
        }

        /// <summary>
        /// Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        /// Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        /// Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            Tracer.WriteTraceMethodLine();
            base.NotifyShutdown();
        }

        /// <summary>
        /// Implements the OnConnection method of the IDTExtensibility2 interface.
        /// Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        /// Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        /// Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        /// Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        void IDTExtensibility2.OnConnection(object application,
          ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                Tracer.WriteTraceMethodLine($"ConnectMode = {connectMode}");
                Application = application as OneNote.Application;
                COMAddIn = addInInst as Office.COMAddIn;
                ConnectMode = connectMode;
                RibbonEventHandler = new UI.RibbonController();

                const string Format = "{0,-11} = {1}";
                Tracer.WriteDebugLine(Format, "ProgId", COMAddIn.ProgId);
                Tracer.WriteDebugLine(Format, "Guid", COMAddIn.Guid);
                Tracer.WriteDebugLine(Format, "Connect", COMAddIn.Connect.ToString());
                Tracer.WriteDebugLine(Format, "Description", COMAddIn.Description);
            }
            catch (Exception e)
            {
                Utils.ExceptionHandler.HandleException(e);
            }
        }

        /// <summary>
        /// Implements the OnDisconnection method of the IDTExtensibility2 interface.
        /// Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        /// Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        /// Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        void IDTExtensibility2.OnDisconnection(
          ext_DisconnectMode disconnectMode, ref Array custom)
        {
            Tracer.WriteTraceMethodLine($"DisconnectMode = {disconnectMode}");

            // If add-in was disconnected with the COM Add-ins dialog, need to manually
            // fire OnBeginShutdown event
            if (disconnectMode != ext_DisconnectMode.ext_dm_HostShutdown)
            {
                ((IDTExtensibility2)this).OnBeginShutdown(ref custom);
            }
        }

        /// <summary>
        /// Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        /// Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        /// Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {
                Tracer.WriteTraceMethodLine();
                base.NotifyStartup();
            }
            catch (Exception e)
            {
                Utils.ExceptionHandler.HandleException(e);
            }
        }

        #endregion

        #region IRibbonExtensibility Member

        /// <summary>
        /// Loads the XML markup to customize the Ribbon UI for this add-in.
        /// </summary>
        /// <param name="ribbonID">The ID for the RibbonX UI</param>
        /// <returns>string</returns>
        string Office.IRibbonExtensibility.GetCustomUI(string ribbonID)
        {
            Tracer.WriteTraceMethodLine($"RibbonID = {ribbonID}");
            return Properties.Resources.UI_CustomUI_v16_0_xml;
        }

        #endregion
    }
}
