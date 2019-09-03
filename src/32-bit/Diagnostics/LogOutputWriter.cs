using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace MyJournal.Notebook.Diagnostics
{
    class LogOutputWriter : TraceOutputWriter
    {
        internal LogOutputWriter(
          FileStream logFileStream, TraceSwitch traceSwitch) : base(traceSwitch)
        {
            DoWriteLine = LogWriteLine;

            Trace.AutoFlush = true;
            var writer = new StreamWriter(logFileStream, Encoding.Unicode);
            Trace.Listeners.Add(new TextWriterTraceListener(writer));
            Trace.Listeners.Remove("Default");

            CallingMethod = GetDefaultMethodBase;

            // Write log entry header
            LogWriteLine("{0} {1}",
              new string('#', 7), DateTime.Now.ToLongDateString());

            base.WriteTraceTypeLine(
              "{0} = {1}",
              Tracer.TRACESWITCH_LEVEL_PROPERTY,
              traceSwitch.Level);

            CallingMethod = GetTraceMethodBase;
        }

        void LogWriteLine(string format, params object[] args)
        {
            Trace.Write(DateTime.Now.ToString("hh:mm:ss.fff tt "));
            Trace.WriteLine(string.Format(format, args));
        }
    }
}
