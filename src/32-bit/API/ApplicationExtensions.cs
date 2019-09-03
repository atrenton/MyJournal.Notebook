using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Win32;
using MyJournal.Notebook.Diagnostics;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.API
{
    /// <summary>
    /// Extends the OneNote 2010 Application Interface.
    /// </summary>
    static class ApplicationExtensions
    {
        const OneNote.CreateFileType
          ONE_CFT_FOLDER = OneNote.CreateFileType.cftFolder,
          ONE_CFT_NOTEBOOK = OneNote.CreateFileType.cftNotebook,
          ONE_CFT_SECTION = OneNote.CreateFileType.cftSection;

        const OneNote.HierarchyScope
          ONE_HS_CHILDREN = OneNote.HierarchyScope.hsChildren,
          ONE_HS_NOTEBOOKS = OneNote.HierarchyScope.hsNotebooks,
          ONE_HS_PAGES = OneNote.HierarchyScope.hsPages,
          ONE_HS_SECTIONS = OneNote.HierarchyScope.hsSections;

        const OneNote.SpecialLocation ONE_SL_DEFAULT_NOTEBOOK_FOLDER =
          OneNote.SpecialLocation.slDefaultNotebookFolder;

        const OneNote.XMLSchema ONE_XML_SCHEMA =
          OneNote.XMLSchema.xs2010;

        /// <summary>
        /// XML document wrapper for OneNote Application.CreateNewPage method.
        /// </summary>
        /// <param name="application">A reference to the OneNote Application
        /// Interface.</param>
        /// <param name="sectionId">A string containing the OneNote Section ID in
        /// which the new page is created.</param>
        /// <param name="pageId">(Output parameter) A string containing the OneNote
        /// ID for the new Page.</param>
        /// <returns>An XML document containing the new blank page with title.</returns>
        internal static XDocument CreatePage(this OneNote.IApplication application,
          string sectionId, out string pageId)
        {
            application.CreateNewPage(sectionId, out pageId,
              OneNote.NewPageStyle.npsBlankPageWithTitle);

            var xmlDocument = GetPage(application, pageId);

            // Take snapshot of OneNote XML for page style:
            // OneNote.NewPageStyle.npsBlankPageWithTitle
            // Utils.WinHelper.SendXmlToClipboard(xmlDocument.ToString());

            return xmlDocument;
        }

        /// <summary>
        /// XML document wrapper for OneNote Application.GetPageContent method.
        /// </summary>
        /// <param name="application">A reference to the OneNote Application
        /// Interface.</param>
        /// <param name="pageId">A string containing the OneNote Page ID to be
        /// retrieved.</param>
        /// <returns>An XML document containing the page content.</returns>
        internal static XDocument GetPage(this OneNote.IApplication application,
          string pageId)
        {
            string xml;
            application.GetPageContent(pageId, out xml, OneNote.PageInfo.piAll,
              ONE_XML_SCHEMA);
            return XDocument.Parse(xml);
        }

        static string GetObjectId(this OneNote.IApplication application,
          string parentId, OneNote.HierarchyScope scope, string objectName)
        {
            Tracer.WriteTraceMethodLine("Name = {0}, Parent ID = {1}", objectName,
              parentId);

            string xml;
            application.GetHierarchy(parentId, scope, out xml, ONE_XML_SCHEMA);

            // Take snapshot of OneNote XML Hierarchy
            //Utils.WinHelper.SendXmlToClipboard(xml);

            var doc = XDocument.Parse(xml);
            var nodeName = string.Empty;

            switch (scope)
            {
                case ONE_HS_PAGES: nodeName = "Page"; break;
                case ONE_HS_SECTIONS: nodeName = "Section"; break;
                case ONE_HS_CHILDREN: nodeName = "SectionGroup"; break;
                case ONE_HS_NOTEBOOKS: nodeName = "Notebook"; break;
                default: return null;
            }

            // Each hierarchy scope element has a name attribute
            // E.G., For a Page element, it's name attribute contains the page title
            var node = doc.Descendants(application.GetXmlNamespace() + nodeName)
                .Where(n => n.Attribute("name").Value == objectName).FirstOrDefault();

            return node?.Attribute("ID").Value;
        }

        /// <summary>
        /// Returns OneNote execututable file path.
        /// </summary>
        internal static string GetExeFilePath()
        {
            const string Root = "Path";
            string exeFilePath = null;

            foreach (var version in s_office_version_list)
            {
                var subKey = string.Format(INSTALL_SUBKEY, version);
                using (var k = Registry.LocalMachine.OpenSubKey(subKey))
                {
                    if (k != null && k.GetValue(Root) != null)
                    {
                        var dir = k.GetValue(Root).ToString();
                        exeFilePath = Path.Combine(dir, ONENOTE_EXE);
                        break;
                    }
                }
            }
            return exeFilePath;
        }

        internal static string GetNotebookId(this OneNote.IApplication application,
          string notebookName)
        {
            string objId, path, s = string.Empty;
            application.GetSpecialLocation(ONE_SL_DEFAULT_NOTEBOOK_FOLDER, out path);
            var notebookPath = Path.Combine(path, notebookName);
            application.OpenHierarchy(notebookPath, s, out objId, ONE_CFT_NOTEBOOK);
            Tracer.WriteTraceMethodLine("Name = {0}, ID = {1}", notebookName, objId);

            return objId;
        }

        internal static string GetSectionGroupId(this OneNote.IApplication application,
          string groupName, string notebookId)
        {
            string objId;
            application.OpenHierarchy(groupName, notebookId, out objId, ONE_CFT_FOLDER);
            Tracer.WriteTraceMethodLine("Name = {0}, ID = {1}", groupName, objId);

            return objId;
        }

        internal static string GetSectionId(this OneNote.IApplication application,
          string sectionName, string groupId)
        {
            string fileName = sectionName + ".one", objId;
            application.OpenHierarchy(fileName, groupId, out objId, ONE_CFT_SECTION);
            Tracer.WriteTraceMethodLine("Name = {0}, ID = {1}", sectionName, objId);

            return objId;
        }

        internal static string GetPageId(this OneNote.IApplication application,
          string pageName, string sectionId)
        {
            var objId = application.GetObjectId(sectionId, ONE_HS_PAGES, pageName);
            Tracer.WriteTraceMethodLine("Name = {0}, ID = {1}", pageName, objId);
            return objId;
        }

        internal static XNamespace GetXmlNamespace(this OneNote.IApplication 
          application)
        {
            if (s_namespace == null)
            {
                string xml = null;
                lock (s_syncNamespace)
                {
                    if (s_namespace == null)
                    {
                        // Set OneNote XML Namespace value
                        application.GetHierarchy(null, ONE_HS_NOTEBOOKS, out xml,
                          ONE_XML_SCHEMA);
                        s_namespace = XDocument.Parse(xml).Root.Name.Namespace;
                    }
                }
            }
            return s_namespace;
        }

        /// <summary>
        /// XML document wrapper for OneNote Application.UpdatePageContent method.
        /// </summary>
        /// <param name="application">A reference to the OneNote Application
        /// Interface.</param>
        /// <param name="document">A XML document that contains the updated Page
        /// content.</param>
        /// <param name="lastModified">A DateTime value that must match the Page
        /// lastModifiedTime attribute value.</param>
        internal static void UpdatePage(this OneNote.IApplication application,
          XDocument document, DateTime lastModified)
        {
            var one = application.GetXmlNamespace();
            if (document.Root.Name != (one + "Page"))
                throw new ArgumentException(
                  "Document root element name != \"one:Page\"", nameof(document));

            var xml = document.ToString();
            // Take snapshot of OneNote XML before updating page content
            // Utils.WinHelper.SendXmlToClipboard(xml);

            application.UpdatePageContent(xml, lastModified, ONE_XML_SCHEMA);
        }

        static volatile XNamespace s_namespace;
        static readonly object s_syncNamespace = new object();

        // Office 2016, 2013 and 2010 version numbers
        static readonly int[] s_office_version_list = { 16, 15, 14 };
        const string
          INSTALL_SUBKEY = @"SOFTWARE\Microsoft\Office\{0}.0\OneNote\InstallRoot",
          ONENOTE_EXE = "onenote.exe";
    }
}
