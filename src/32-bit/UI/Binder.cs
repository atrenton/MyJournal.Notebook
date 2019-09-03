using System;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Templates;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.UI
{
    class Binder : IDisposable
    {
        bool _disposed;

        #region Properties

        /// <summary>
        /// Page settings data model.
        /// </summary>
        internal Config.IPageSettingsModel PageSettings
        {
            get;
            private set;
        }

        /// <summary>
        /// Page template presentation logic.
        /// </summary>
        internal Templates.IPageTemplatePresenter PageTemplate
        {
            get;
            private set;
        }

        #endregion

        /// <summary>
        /// Provides a one-way binding between the page settings data model
        /// (source) and page template presentation logic (target).
        /// </summary>
        internal Binder(OneNote.IApplication application)
        {
            PageSettings = PageSettingsDataSource.Load();
            BindTemplate(application);
        }
        ~Binder()
        {
            Dispose(false);
        }

        internal void CreateJournalPage()
        {
            Tracer.WriteTraceMethodLine();
            PageTemplate.CreateNewPage(PageSettings, EventArgs.Empty);
        }

        #region IDisposable Member

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing) // dispose of managed resources
                {
                    Tracer.WriteTraceMethodLine();
                    PageSettingsDataSource.Save(PageSettings);
                    PageSettings = null;
                    PageTemplate = null;
                }
                _disposed = true;
            }
        }

        #endregion

        internal void Rebind(OneNote.IApplication application) =>
          BindTemplate(application);

        private void BindTemplate(OneNote.IApplication application)
        {
            var template = PageTemplate;
            if (template != null) UnbindTemplate(template);

            var model = PageSettings;
            template = TemplateFactory.CreatePageTemplate(application);
            model.ColorChanged += template.ChangeColor;
            model.RuleLinesHorizontalColorChanged +=
              template.ChangeRuleLinesHorizontalColor;
            model.RuleLinesHorizontalSpacingChanged +=
              template.ChangeRuleLinesHorizontalSpacing;
            model.RuleLinesMarginColorChanged +=
              template.ChangeRuleLinesMarginColor;
            model.RuleLinesVisibleChanged += template.ChangeRuleLinesVisible;
            model.TitleChanged += template.ChangeTitle;
            PageTemplate = template;
            Tracer.WriteTraceMethodLine();
        }

        private void UnbindTemplate(Templates.IPageTemplatePresenter template)
        {
            Tracer.WriteTraceMethodLine();
            var model = PageSettings;
            model.ColorChanged -= template.ChangeColor;
            model.RuleLinesHorizontalColorChanged -=
              template.ChangeRuleLinesHorizontalColor;
            model.RuleLinesHorizontalSpacingChanged -=
              template.ChangeRuleLinesHorizontalSpacing;
            model.RuleLinesMarginColorChanged -=
              template.ChangeRuleLinesMarginColor;
            model.RuleLinesVisibleChanged -= template.ChangeRuleLinesVisible;
            model.TitleChanged -= template.ChangeTitle;
        }
    }
}
