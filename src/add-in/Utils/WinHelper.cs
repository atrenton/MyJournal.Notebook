using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyJournal.Notebook.Diagnostics;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Utils
{
    static class WinHelper
    {
        static readonly string s_caption;
        static readonly bool s_isOneNoteRunning;

        internal class Win32Window : IWin32Window
        {
            internal Win32Window(IntPtr handle)
            {
                Handle = handle;
            }

            internal Win32Window(OneNote.Window context)
              : this((IntPtr)context.WindowHandle)
            {
            }

            #region IWin32Window Methods

            public IntPtr Handle
            {
                get;
                private set;
            }

            #endregion
        }

        [DllImport("user32.dll", SetLastError = false)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        static extern int SetForegroundWindow(IntPtr hWnd);

#pragma warning disable CA1810
        static WinHelper()
        {
            s_caption = Component.AssemblyInfo.Title;
            s_isOneNoteRunning = (Process.GetProcessesByName("OneNote").Length > 0);
        }
#pragma warning restore CA1810

        #region Error Messages

        internal static void DisplayError(string text) =>
            DisplayModal(GetWindowOwner(), text, MessageBoxIcon.Error);

        internal static async Task
        DisplayErrorAsync(OneNote.Window owner, string[] text)
        {
            var multiLineText = string.Join(Environment.NewLine, text);
            await DisplayErrorAsync(owner, multiLineText).ConfigureAwait(false);
        }

        internal static async Task
        DisplayErrorAsync(OneNote.Window owner, string text)
        {
            const MessageBoxIcon Icon = MessageBoxIcon.Error;
            var parent = new Win32Window(owner);
            await DisplayAsync(parent, text, Icon).ConfigureAwait(false);
        }

        #endregion

        #region Informational Messages

        internal static void DisplayInfo(string text) =>
            DisplayModal(GetWindowOwner(), text, MessageBoxIcon.Information);

        internal static async Task
        DisplayInfoAsync(OneNote.Window owner, string[] text)
        {
            var multiLineText = string.Join(Environment.NewLine, text);
            await DisplayInfoAsync(owner, multiLineText).ConfigureAwait(false);
        }

        internal static async Task
        DisplayInfoAsync(OneNote.Window owner, string text)
        {
            const MessageBoxIcon Icon = MessageBoxIcon.Information;
            var parent = new Win32Window(owner);
            await DisplayAsync(parent, text, Icon).ConfigureAwait(false);
        }

        #endregion

        #region Warning Messages

        internal static void DisplayWarning(string text) =>
            DisplayModal(GetWindowOwner(), text, MessageBoxIcon.Warning);

        internal static async Task
        DisplayWarningAsync(OneNote.Window owner, string[] text)
        {
            var multiLineText = string.Join(Environment.NewLine, text);
            await DisplayWarningAsync(owner, multiLineText).ConfigureAwait(false);
        }

        internal static async Task
        DisplayWarningAsync(OneNote.Window owner, string text)
        {
            const MessageBoxIcon Icon = MessageBoxIcon.Warning;
            var parent = new Win32Window(owner);
            await DisplayAsync(parent, text, Icon).ConfigureAwait(false);
        }

        #endregion

        static async Task DisplayAsync(string text, MessageBoxIcon ico) =>
          await DisplayAsync(GetWindowOwner(), text, ico).ConfigureAwait(false);

        static async Task
        DisplayAsync(IWin32Window owner, string text, MessageBoxIcon ico)
        {
            Action action = () => DisplayModal(owner, text, ico);
            await Task.Run(action).ConfigureAwait(false);
        }

        static void DisplayModal(IWin32Window owner, string text, MessageBoxIcon ico)
        {
            try
            {
                MessageBox.Show(owner, text, s_caption, MessageBoxButtons.OK, ico);
            }
            catch (Exception e)
            {
                ExceptionHandler.HandleException(e);
            }
        }

        static IWin32Window GetWindowOwner()
        {
            return (s_isOneNoteRunning) ?
              new Win32Window(GetForegroundWindow()) :
              new Win32Window(GetDesktopWindow());
        }

        internal static void SendXmlToClipboard(string xml)
        {
            var t = new Thread(WinHelper.Clipboard_SetText);
            t.SetApartmentState(ApartmentState.STA);
            t.Start(xml);
        }

        internal static void SetFocus(ulong hWndFocus)
        {
            var format = "hWndCurrent = 0x{0:X8}, hWndFocus = 0x{1:X8}";
            var hWndCurrent = (ulong)GetForegroundWindow();
            Tracer.WriteTraceMethodLine(format, hWndCurrent, hWndFocus);
            if (hWndFocus != 0)
            {
                var hResult = SetForegroundWindow((IntPtr)hWndFocus);
            }
        }

        static void Clipboard_SetText(object data)
        {
            try
            {
                var text = data as string;
                Clipboard.SetText(text);
                text = "The XML has been copied to the clipboard.";
                DisplayModal(GetWindowOwner(), text, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                ExceptionHandler.HandleException(e);
            }
        }
    }
}
