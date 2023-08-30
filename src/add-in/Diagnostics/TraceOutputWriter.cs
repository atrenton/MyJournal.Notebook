using System.Diagnostics;
using System.Reflection;

namespace MyJournal.Notebook.Diagnostics
{
    class TraceOutputWriter : DefaultOutputWriter
    {
        TraceSwitch _switch;

        internal TraceOutputWriter(TraceSwitch traceSwitch)
        {
            _switch = traceSwitch;
            CallingMethod = GetTraceMethodBase;
            DoWriteLine = TraceWriteLine;
        }

        /// <summary>
        /// Logs a formatted debugging message.
        /// Message is output when TraceLevel = Verbose.
        /// </summary>
        internal override void WriteDebugLine(object text, params object[] args)
        {
            if (_switch.TraceVerbose) base.WriteDebugLine(text, args);
        }

        /// <summary>
        /// Logs a formatted trace message.
        /// Message is output when TraceLevel = Verbose.
        /// </summary>
        internal override void WriteTraceLine(object text, params object[] args)
        {
            if (_switch.TraceInfo) base.WriteTraceLine(text, args);
        }

        /// <summary>
        /// Logs the name of the calling method.
        /// Message is output when trace level = Verbose.
        /// </summary>
        internal override void WriteTraceMethodLine()
        {
            if (_switch.TraceInfo) base.WriteTraceMethodLine();
        }

        /// <summary>
        /// Logs the name of the calling method and appends a formatted arg list.
        /// Message is output when trace level = Verbose.
        /// </summary>
        internal override void WriteTraceMethodLine(object text, params object[] args)
        {
            if (_switch.TraceInfo) base.WriteTraceMethodLine(text, args);
        }

        /// <summary>
        /// Logs the name of the calling type.
        /// Message is output when trace level = Verbose.
        /// </summary>
        internal override void WriteTraceTypeLine()
        {
            if (_switch.TraceInfo) base.WriteTraceTypeLine();
        }

        /// <summary>
        /// Logs the name of the calling object type and appends a formatted arg list.
        /// Message is output when trace level = Verbose.
        /// </summary>
        internal override void WriteTraceTypeLine(object text, params object[] args)
        {
            if (_switch.TraceInfo) base.WriteTraceTypeLine(text, args);
        }

        /// <summary>
        /// Logs a formatted informational message.
        /// Message is output when TraceLevel = Info.
        internal override void WriteInfoLine(object text, params object[] args)
        {
            if (_switch.TraceInfo) base.WriteInfoLine(text, args);
        }

        /// <summary>
        /// Logs a formatted informational message.
        /// Message is output when TraceLevel = Warning.
        internal override void WriteWarnLine(object text, params object[] args)
        {
            if (_switch.TraceWarning) base.WriteWarnLine(text, args);
        }

        /// <summary>
        /// Logs a formatted error message.
        /// Message is output when TraceLevel = Error.
        /// </summary>
        internal override void WriteErrorLine(object text, params object[] args)
        {
            if (_switch.TraceError) base.WriteErrorLine(text, args);
        }

        internal override void WriteDittoLine(object text, params object[] args)
        {
            if (_switch.TraceInfo) base.WriteDittoLine(text, args);
        }

        void TraceWriteLine(string format, params object[] args)
        {
            Trace.WriteLine(string.Format(format, args));
        }

        protected internal MethodBase GetTraceMethodBase() =>
            new StackFrame(4, false).GetMethod();
    }
}
