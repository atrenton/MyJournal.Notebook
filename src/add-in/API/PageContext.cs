using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Utils;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.API
{
    class PageContext
    {
        private OneNote.IApplication _Application { get; set; }
        private XDocument _CurrentPage
        {
            get
            {
                return HasPageId() ? _Application.GetPage(PageId) : null;
            }
        }

        #region Properties

        internal string NotebookId { get; private set; }
        internal string SectionGroupId { get; private set; }
        internal string SectionId { get; private set; }
        internal string PageId { get; private set; }
        internal string PageName { get; private set; }

        #endregion

        #region Constructors

        private PageContext(OneNote.IApplication application)
        {
            _Application = application;
        }

        internal PageContext(OneNote.IApplication application, PageTitleEnum pageTitle)
          : this(application)
        {
            var today = DateTime.Today;
            var sectionGroup = today.Year.ToString();
            var section = today.ToString("MMMM");

            NotebookId = _Application.GetNotebookId(Component.FriendlyName);
            SectionGroupId = _Application.GetSectionGroupId(sectionGroup, NotebookId);
            SectionId = _Application.GetSectionId(section, SectionGroupId);
            PageName = today.Format(pageTitle);
            PageId = _Application.GetPageId(PageName, SectionId);
        }

        internal static PageContext CurrentWindow(OneNote.IApplication application)
        {
            var context = new PageContext(application);
            var currentWindow = application.Windows.CurrentWindow;
            context.NotebookId = currentWindow.CurrentNotebookId;
            context.SectionGroupId = currentWindow.CurrentSectionGroupId;
            context.SectionId = currentWindow.CurrentSectionId;
            context.PageId = currentWindow.CurrentPageId;
            context.PageName = context._CurrentPage?.Root.Attribute("name").Value;
            return context;
        }

        #endregion

        internal XDocument CreateNewPage()
        {
            var page = _CurrentPage;
            if (page == null)
            {
                string pageId;
                page = _Application.CreatePage(SectionId, out pageId);
                page.Root.SetAttributeValue(STATIONERY_NAME, Component.FriendlyName);
                PageId = pageId;
                Tracer.WriteTraceMethodLine("Name = {0}, ID = {1}", PageName, PageId);
            }
            return page;
        }

        bool HasPageId() => !string.IsNullOrEmpty(PageId);

        /// <summary>
        /// Validates page context and outputs current page content, if valid.
        /// </summary>
        /// <param name="page">My Journal notebook page XML content</param>
        /// <returns>true if current page context is in My Journal notebook</returns>
        internal bool IsMyJournalNotebook(out XDocument page)
        {
            var currentPage = _CurrentPage;

            var name = currentPage?.Root.Attribute(STATIONERY_NAME);
            page = (name?.Value == Component.FriendlyName) ? currentPage : null;

            return (page != null);
        }

        /// <summary>
        /// Navigates to the page specified by the PageId property.
        /// </summary>
        internal void NavigateToPage()
        {
            if (HasPageId())
            {
                _Application.NavigateTo(PageId, null, false);
                var windowHandle = _Application.Windows.CurrentWindow.WindowHandle;
                Tracer.WriteTraceMethodLine("WindowHandle = 0x{0:X8}", windowHandle);
                Utils.WinHelper.SetFocus(windowHandle);
            }
        }

        internal bool PageNotFound() => string.IsNullOrEmpty(PageId);

        internal void SaveCurrentPageToDisk()
        {
            var fileName = PageName;
            var invalidCharList = Path.GetInvalidFileNameChars().ToList<char>();
            invalidCharList.ForEach(ch => fileName = fileName.Replace(ch, '-'));

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var dir = Path.Combine(desktop, "MyJournal");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var filePath = Path.Combine(dir, fileName + ".xml");
            Tracer.WriteTraceMethodLine("FilePath = {0}", filePath);

            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                _CurrentPage.Save(fs);
            }
        }

        internal void SendCurrentPageToClipboard()
        {
            var builder = new StringBuilder();
            using (var writer = new StringWriter(builder))
            {
                _CurrentPage.Save(writer);
            }
            WinHelper.SendXmlToClipboard(builder.ToString());
        }

        /// <summary>
        /// Saves updated page content to My Journal notebook.
        /// </summary>
        /// <param name="updatedPageContent">page content in XML format</param>
        /// <param name="lastModified">timestamp used to enforce data integrity;
        /// if null, the original timestamp value contained in the page is used</param>
        internal void UpdateMyJournal(XDocument updatedPageContent,
          DateTime? lastModified = null)
        {
            var lastModifiedTime = lastModified ??
                    (DateTime)updatedPageContent.Root.Attribute("lastModifiedTime");

            _Application.UpdatePage(updatedPageContent, lastModifiedTime);
        }

        const string STATIONERY_NAME = "stationeryName";
    }
}
