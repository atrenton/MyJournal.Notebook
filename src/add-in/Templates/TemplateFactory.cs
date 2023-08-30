using System;
using System.Reflection;
using MyJournal.Notebook.Diagnostics;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Templates
{
    static class TemplateFactory
    {
        static readonly string s_baseTemplateName;
        static readonly int s_baseIndex;

        #region Properties

        /// <summary>
        /// Gets the collection of page template names.
        /// </summary>
        internal static object[] Items =>
          new[] { "Default", "Bullet", "OrdinalDay", "HappyDay", "Retro" };

        internal static Type PageTemplateType
        {
            get
            {
                var value = Properties.Settings.Default.PageTemplate;
                var typeName = s_baseTemplateName.Insert(s_baseIndex, value);
                return Type.GetType(typeName) ?? typeof(DefaultPageTemplate);
            }
        }

        #endregion

#pragma warning disable CA1810
        static TemplateFactory()
        {
            const StringComparison Compare = StringComparison.CurrentCulture;
            s_baseTemplateName = typeof(PageTemplate).FullName;
            s_baseIndex = s_baseTemplateName.IndexOf("PageTemplate", Compare);
        }
#pragma warning restore CA1810

        /// <summary>
        /// Creates a page template instance.
        /// </summary>
        /// <param name="application">OneNote application reference</param>
        /// <returns>Reference to page template presentation logic</returns>
        internal static IPageTemplatePresenter
          CreatePageTemplate(OneNote.IApplication application)
        {
            var type = PageTemplateType;
            Tracer.WriteTraceMethodLine("Template = {0}", type.FullName);

            var flags = BindingFlags.NonPublic | BindingFlags.Instance;

            return (IPageTemplatePresenter)
              Activator.CreateInstance(
                  type, flags, null, new[] { application }, null);
        }
    }
}
