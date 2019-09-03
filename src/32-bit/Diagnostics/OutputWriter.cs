using System;
using System.Reflection;

namespace MyJournal.Notebook.Diagnostics
{
    delegate void WriteAction(string format, params object[] args);

    abstract class OutputWriter
    {
        internal OutputWriter() { }

        protected Func<MethodBase> CallingMethod { get; set; }
        protected WriteAction DoWriteLine { get; set; }

        internal abstract void WriteDebugLine(object text, params object[] args);

        internal abstract void WriteTraceLine(object text, params object[] args);

        internal abstract void WriteTraceMethodLine();

        internal abstract void WriteTraceMethodLine(
          object text, params object[] args);

        internal abstract void WriteTraceTypeLine();

        internal abstract void WriteTraceTypeLine(
          object text, params object[] args);

        internal abstract void WriteInfoLine(object text, params object[] args);

        internal abstract void WriteWarnLine(object text, params object[] args);

        internal abstract void WriteErrorLine(object text, params object[] args);

        internal abstract void WriteDittoLine(object text, params object[] args);

        internal abstract void WriteLine(object text, params object[] args);
    }
}
