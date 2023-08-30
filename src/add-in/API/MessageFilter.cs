using System;
using System.Runtime.InteropServices;
using System.Threading;
using MyJournal.Notebook.Diagnostics;

namespace MyJournal.Notebook.API
{
    internal sealed class MessageFilter : IMessageFilter, IDisposable
    {
        [DllImport("ole32.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern int CoRegisterMessageFilter(IMessageFilter newFilter,
            out IMessageFilter oldFilter);

        private bool _disposed;
        private bool _isRegistered;
        private IMessageFilter _oldFilter;

        /// <summary>
        /// Implements a COM Message Filter.
        /// </summary>
        public MessageFilter()
        {
            Register();
        }

        private void Register()
        {
            Thread.CurrentThread.SetApartmentState(ApartmentState.STA);

            var result = CoRegisterMessageFilter(this, out _oldFilter);
            if (result != 0)
            {
                throw new COMException("CoRegisterMessageFilter failed", result);
            }
            _isRegistered = true;
        }

        private void Revoke()
        {
            if (_isRegistered)
            {
                IMessageFilter revokedFilter;
                var hr = CoRegisterMessageFilter(_oldFilter, out revokedFilter);
                _oldFilter = null;
                _isRegistered = false;
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing) // dispose of unmanaged resources
            {
                Tracer.WriteTraceMethodLine();
                Revoke();
            }
            _disposed = true;
        }

        ~MessageFilter() => Dispose();

        #endregion

        #region IMessageFilter Members

        int IMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller,
            int dwTickCount, IntPtr lpInterfaceInfo) => (int)SERVERCALL.ISHANDLED;


        // REF: https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-imessagefilter-retryrejectedcall
        int IMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount,
            int dwRejectType)
        {
            if (dwRejectType == (int)SERVERCALL.REJECTED ||
                dwRejectType == (int)SERVERCALL.RETRYLATER)
            {
                var rejectReason = (dwRejectType < 2) ?
                    "The call was rejected" : "The application is busy";

                var millis = 250;
                var msg = $"{rejectReason}, sleeping . . .";

                Tracer.WriteInfoLine("{0}: {1}", GetType().FullName, msg);

                return millis; // wait and try again
            }
            Tracer.WriteWarnLine("Got dwRejectType = {0}", dwRejectType);
            return -1; // cancel call
        }

        int IMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount,
            int dwPendingType) => (int)PENDINGMSG.WAITDEFPROCESS;

        #endregion
    }

    enum SERVERCALL
    {
        ISHANDLED = 0,
        REJECTED = 1,
        RETRYLATER = 2
    }

    enum PENDINGMSG
    {
        CANCELCALL = 0,
        WAITNOPROCESS = 1,
        WAITDEFPROCESS = 2
    }
}
