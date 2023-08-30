using System;
using System.Runtime.InteropServices;

namespace MyJournal.Notebook.API
{
    // REF: https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-imessagefilter

    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller,
            int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount,
            int dwPendingType);
    }
}
