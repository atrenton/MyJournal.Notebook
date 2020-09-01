using System;
using System.Windows.Forms;

namespace MyJournal.Notebook.Diagnostics
{
    static class Tracer
    {
        internal static OutputWriter Instance { get; private set; }

        internal static void Initialize()
        {
            var displayName = Component.FriendlyName;
            try
            {
                var factory = new OutputWriterFactory
                {
                    Description = Component.Description,
                    DisplayName = displayName,

                    OutputWriterTypeName =
                    Component.AppSettings[OUTPUTWRITER_TYPE_NAME_PROPERTY],

                    TraceSwitchLevel =
                    Component.AppSettings[TRACESWITCH_LEVEL_PROPERTY] ?? "Off"
                };
                Instance = factory.CreateInstance();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error initializing Add-in:\r\n" +
                                e.Message, displayName, MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        internal static void WriteDebugLine(object text, params object[] args) =>
            Instance?.WriteDebugLine(text, args);

        internal static void WriteTraceLine(object text, params object[] args) =>
            Instance?.WriteTraceLine(text, args);

        internal static void WriteTraceMethodLine() =>
            Instance?.WriteTraceMethodLine();

        internal static void WriteTraceMethodLine(
          object text, params object[] args) =>
            Instance?.WriteTraceMethodLine(text, args);

        internal static void WriteTraceTypeLine() =>
            Instance?.WriteTraceTypeLine();

        internal static void WriteTraceTypeLine(
          object text, params object[] args) =>
            Instance?.WriteTraceTypeLine(text, args);

        internal static void WriteInfoLine(object text, params object[] args) =>
            Instance?.WriteInfoLine(text, args);

        internal static void WriteWarnLine(object text, params object[] args) =>
            Instance?.WriteWarnLine(text, args);

        internal static void WriteErrorLine(object text, params object[] args) =>
            Instance?.WriteErrorLine(text, args);

        internal static void WriteDittoLine(object text, params object[] args) =>
            Instance?.WriteDittoLine(text, args);

        internal static void WriteLine(object text, params object[] args) =>
            Instance?.WriteLine(text, args);

        internal static void WriteStackTrace(Exception e)
        {
            if (Instance != null)
            {
                var nl = new[] { Environment.NewLine };

                var lines = e.StackTrace.Split(
                    nl, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines) Instance.WriteErrorLine(line);
            }
        }

        internal const string
          OUTPUTWRITER_TYPE_NAME_PROPERTY = "Diagnostics.OutputWriter.Type.Name",
          TRACESWITCH_LEVEL_PROPERTY = "Diagnostics.TraceSwitch.Level";
    }
}
