using System;
using System.Diagnostics;
using System.Reflection;

namespace MyJournal.Notebook.Diagnostics
{
    class DefaultOutputWriter : OutputWriter
    {
        const string
            DEBUG = "[DEBUG] ",
            DITTO = "[  \"  ] ",
            ERROR = "[ERROR] ",
            INFO = "[INFO ] ",
            TRACE = "[TRACE] ",
            WARN = "[WARN ] ";

        internal DefaultOutputWriter()
        {
            CallingMethod = GetDefaultMethodBase;
            DoWriteLine = Console.WriteLine;
        }

        internal override void WriteDebugLine(object text, params object[] args)
        {
            DoWriteLine(DEBUG + text, args);
        }

        internal override void WriteTraceLine(object text, params object[] args)
        {
            DoWriteLine(TRACE + text, args);
        }

        internal override void WriteTraceMethodLine()
        {
            var mb = CallingMethod();
            WriteTraceLine("{0}.{1}()", mb.DeclaringType.FullName, mb.Name);
        }

        internal override void WriteTraceMethodLine(
          object text, params object[] args)
        {
            var mb = CallingMethod();
            WriteTraceLine("{0}.{1}() =>", mb.DeclaringType.FullName, mb.Name);
            WriteDittoLine(text, args);
        }

        internal override void WriteTraceTypeLine()
        {
            WriteTraceLine("{0} created", CallingMethod().DeclaringType.FullName);
        }

        internal override void WriteTraceTypeLine(
          object text, params object[] args)
        {
            WriteTraceLine("{0} created => ",
                CallingMethod().DeclaringType.FullName);

            WriteDittoLine(text, args);
        }

        internal override void WriteInfoLine(object text, params object[] args)
        {
            DoWriteLine(INFO + text, args);
        }

        internal override void WriteWarnLine(object text, params object[] args)
        {
            DoWriteLine(WARN + text, args);
        }

        internal override void WriteErrorLine(object text, params object[] args)
        {
            DoWriteLine(ERROR + text, args);
        }

        internal override void WriteDittoLine(object text, params object[] args)
        {
            DoWriteLine(DITTO + text, args);
        }

        internal override void WriteLine(object text, params object[] args)
        {
            DoWriteLine(text.ToString(), args);
        }

        protected internal MethodBase GetDefaultMethodBase() =>
            new StackFrame(3, false).GetMethod();
    }
}
