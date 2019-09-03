using System;
using System.Diagnostics;
using System.IO;

namespace MyJournal.Notebook.Diagnostics
{
    class OutputWriterFactory
    {
        internal string Description { get; set; }
        internal string DisplayName { get; set; }
        internal string TraceSwitchLevel { get; set; }
        internal string OutputWriterTypeName { get; set; }

        internal OutputWriterFactory() { }

        internal OutputWriter CreateInstance()
        {
            OutputWriter writer = null;
            var ts = new TraceSwitch(DisplayName, Description)
            {
                Level = (TraceLevel)
                    Enum.Parse(typeof(TraceLevel), TraceSwitchLevel, true)
            };

            if (ts.Level != TraceLevel.Off)
            {
                switch (OutputWriterTypeName)
                {
                    case "LogOutputWriter":
                    case "logoutputwriter":
                        var logFileName = DisplayName + ".log";
                        var logFilePath =
                            Path.Combine(Path.GetTempPath(), logFileName);

                        var logFileStream =
                            new FileStream(logFilePath, FileMode.Append);

                        writer = new LogOutputWriter(logFileStream, ts);
                        break;

                    case "TraceOutputWriter":
                    case "traceoutputwriter":
                        writer = new TraceOutputWriter(ts);
                        break;

                    default:
                        writer = new DefaultOutputWriter();
                        break;
                }
            }
            return writer;
        }
    }
}
