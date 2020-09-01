using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Win32;
using MyJournal.Notebook.Diagnostics;
using HRESULT = System.Int32;
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

        private static Dictionary<HRESULT, string> s_errorCodeTable;

        static ApplicationExtensions()
        {
            // Initialize OneNote version lookup table
            s_versionLookup.Add("12", "2007");
            s_versionLookup.Add("14", "2010");
            s_versionLookup.Add("15", "2013");
            s_versionLookup.Add("16", "2016");
        }

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
        /// Returns the OneNote Error Code table.
        /// </summary>
        internal static IReadOnlyDictionary<HRESULT, string> ErrorCodeTable
        {
            get
            {
                if (null == s_errorCodeTable)
                {
                    LoadErrorCodeTable();
                }
                return s_errorCodeTable;
            }
        }

        /// <summary>
        /// Returns OneNote application friendly name.
        /// </summary>
        internal static string FriendlyName
        {
            get
            {
                string friendlyName = null;
                var exeFilePath = GetExeFilePath();
                if (File.Exists(exeFilePath))
                {
                    var path = exeFilePath.Split(Path.DirectorySeparatorChar);
                    var officeProductName = path[path.Length - 2];
                    var match = Regex.Match(officeProductName, @"Office(\d\d)",
                                            RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        var value = match.Groups[1].Value;
                        friendlyName = $"OneNote {s_versionLookup[value]}";
                    }
                }
                return friendlyName;
            }
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

        internal static XNamespace GetXmlNamespace(
          this OneNote.IApplication application)
        {
            if (s_namespace == null)
            {
                string xml = null;
                lock (s_syncApplication)
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
        /// Check if HRESULT value is a OneNote API Error Code.
        /// </summary>
        /// <param name="errorCode">An Exception.HResult property value.</param>
        /// <returns>True if value is a OneNote API error code; false otherwise.</returns>
        internal static bool IsErrorCode(HRESULT errorCode) =>
            ErrorCodeTable.ContainsKey(errorCode);

        internal static bool IsOneNoteInstalled() => File.Exists(GetExeFilePath());

        // REF: https://docs.microsoft.com/en-us/office/client-developer/onenote/error-codes-onenote
        static void LoadErrorCodeTable()
        {
            lock (s_syncApplication)
            {
                if (null == s_errorCodeTable)
                {
                    s_errorCodeTable = new Dictionary<HRESULT, string>(48)
                    {
                        { -2147213312, "The XML is not well-formed." },
                        { -2147213311, "The XML is invalid." },
                        { -2147213310, "The section could not be created." },
                        { -2147213309, "The section could not be opened." },
                        { -2147213308, "The section does not exist." },
                        { -2147213307, "The page does not exist." },
                        { -2147213306, "Local notebooks not supported." },
                        { -2147213305, "The image could not be inserted." },
                        { -2147213304, "The ink could not be inserted." },
                        { -2147213303, "The HTML could not be inserted." },
                        { -2147213302, "The page could not be opened." },
                        { -2147213301, "The section is read-only." },
                        { -2147213300, "The page is read-only." },
                        { -2147213299, "The outline text could not be inserted." },
                        { -2147213298, "The page object does not exist." },
                        { -2147213297, "The binary object does not exist." },
                        { -2147213296, "The last modified date does not match." },
                        { -2147213295, "The section group does not exist." },
                        { -2147213294, "The page does not exist in the section group." },
                        { -2147213293, "There is no active selection." },
                        { -2147213292, "The object does not exist." },
                        { -2147213291, "The notebook does not exist." },
                        { -2147213290, "The file could not be inserted." },
                        { -2147213289, "The name is invalid." },
                        { -2147213288, "The folder (section group) does not exist." },
                        { -2147213287, "The query is invalid." },
                        { -2147213286, "The file already exists." },
                        { -2147213285, "The section is encrypted and locked." },
                        { -2147213284, "The action is disabled by a policy." },
                        { -2147213283, "OneNote has not yet synchronized content." },
                        { -2147213282, "The section is from OneNote 2007 or earlier." },
                        { -2147213281, "The merge operation failed." },
                        { -2147213280, "The XML Schema is invalid." },
                        { -2147213278, "Content loss has occurred (from future versions of OneNote)." },
                        { -2147213277, "The action timed out." },
                        { -2147213276, "Audio recording is in progress." },
                        { -2147213275, "The linked-note state is unknown." },
                        { -2147213274, "No short name exists for the linked note." },
                        { -2147213273, "No friendly name exists for the linked note." },
                        { -2147213272, "The linked note URI is invalid." },
                        { -2147213271, "The linked note thumbnail is invalid." },
                        { -2147213270, "The importation of linked note thumbnail failed." },
                        { -2147213269, "Unread highlighting is disabled for the notebook." },
                        { -2147213268, "The selection is invalid." },
                        { -2147213267, "The conversion failed." },
                        { -2147213266, "Edit failed in the Recycle Bin." },
                        { -2147213264, "A modal dialog is blocking the app." }
                    };
                }
            }
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

        static readonly object s_syncApplication = new object();

        static readonly Hashtable s_versionLookup = new Hashtable(4);

        // Office 2016, 2013 and 2010 version numbers
        static readonly int[] s_office_version_list = { 16, 15, 14 };
        const string
          INSTALL_SUBKEY = @"SOFTWARE\Microsoft\Office\{0}.0\OneNote\InstallRoot",
          ONENOTE_EXE = "onenote.exe";
    }
}
