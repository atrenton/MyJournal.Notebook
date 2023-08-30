using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using MyJournal.Notebook.Diagnostics;

namespace MyJournal.Notebook.Config
{
    internal static class PageSettingsDataSource
    {
        static XmlSerializer s_serializer =
          XmlSerializer.FromTypes(new[] { typeof(PageSettings) })[0];

        static readonly string s_xmlFileName =
          Path.Combine(Component.AppDataPath, "PageSettings.xml");

        static XmlReaderSettings CreateXmlReaderSettings()
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings,
                ValidationType = ValidationType.Schema,
                XmlResolver = null
            };

            var schemaLocation = Properties.Resources.Config_PageSettings_v1_0_xsd;
            var targetNamespace = Properties.Resources.Config_PageSettings_Namespace;
            using (var schema = XmlReader.Create(new StringReader(schemaLocation)))
            {
                settings.Schemas.Add(targetNamespace, schema);
            }
            return settings;
        }

        internal static IPageSettingsModel Load()
        {
            PageSettings ps = null;
            Tracer.WriteTraceMethodLine();
            var setDefaultValues = false;

            if (File.Exists(s_xmlFileName))
            {
                try
                {
                    var settings = CreateXmlReaderSettings();
                    settings.ValidationEventHandler += (sender, e) =>
                    {
                        var severity =
                        e.Severity.ToString().ToLower(CultureInfo.CurrentCulture);

                        Tracer.WriteErrorLine(
                            "A validation {0} has occurred for {1}",
                            severity, s_xmlFileName);

                        Tracer.WriteErrorLine(e.Message);
                        setDefaultValues = true;
                    };

                    using (var input = new StreamReader(s_xmlFileName))
                    {
                        using (var reader = XmlReader.Create(input, settings))
                        {
                            ps = s_serializer.Deserialize(reader) as PageSettings;
                        }
                    }
                }
                catch (InvalidOperationException)
                {
                    Tracer.WriteErrorLine(
                        "PageSettings are invalid; default values loaded");
                    setDefaultValues = true;
                }
                catch (Exception e)
                {
                    Utils.ExceptionHandler.HandleException(e);
                }
            }
            if (ps == null || setDefaultValues) // use default page settings
            {
                ps = new PageSettings(true);
            }
            return ps;
        }

        internal static void Save(IPageSettingsModel pageSettings)
        {
            if (pageSettings.IsModified())
            {
                Tracer.WriteTraceMethodLine();
                try
                {
                    var settings = new XmlWriterSettings
                    {
                        Encoding = Encoding.UTF8,
                        Indent = true
                    };
                    using (var writer = XmlWriter.Create(s_xmlFileName, settings))
                    {
                        s_serializer.Serialize(writer, pageSettings);
                    }
                }
                catch (Exception e)
                {
                    Utils.ExceptionHandler.HandleException(e);
                }
            }
        }
    }
}
