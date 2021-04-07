using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using HRESULT = System.Int32;

using static MyJournal.Notebook.Diagnostics.Tracer;


namespace MyJournal.Notebook.Utils
{
    static class ExceptionHandler
    {
        static readonly List<HRESULT> s_hResultList;

        static ExceptionHandler()
        {
            s_hResultList = new List<HRESULT> {
                RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER,
                SYNCHRONIZING_NOTEBOOK, COR_E_XML
            };
        }

        internal static string FormatHResult(HRESULT hr) =>
            string.Format("HRESULT: 0x{0:X8}", hr);

        internal static void HandleException(object o)
        {
            var e = o as Exception;
            if (e != null)
            {
                WriteErrorLine("{0}: {1}", e.GetType(), e.Message);
                if (e is COMException)
                {
                    var hresult = FormatHResult(e.HResult);
                    if (!e.Message.Contains(hresult)) WriteErrorLine(hresult);
                }
                WriteStackTrace(e);
            }
            else WriteErrorLine(o);
        }

        internal static bool IsTransientError(Exception e)
        {
            if (e != null)
            {
                if (s_hResultList.Contains(e.HResult))
                {
                    WriteWarnLine("{0}: {1}", e.GetType(), e.Message);
                    var hresult = FormatHResult(e.HResult);
                    if (!e.Message.Contains(hresult)) WriteWarnLine(hresult);
                    return true;
                }
            }
            return false;
        }

        internal const HRESULT
            COR_E_XML = unchecked((HRESULT)0x80131940),
            RPC_E_CALL_REJECTED = unchecked((HRESULT)0x80010001),
            RPC_E_SERVERCALL_RETRYLATER = unchecked((HRESULT)0x8001010A),
            SYNCHRONIZING_NOTEBOOK = unchecked((HRESULT)0x8004201D);
    }
}
