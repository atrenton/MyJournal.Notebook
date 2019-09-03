using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;

using OneNote = Microsoft.Office.Interop.OneNote;

using TemplateMethod = System.Action<MyJournal.Notebook.API.PageContext,
  System.Xml.Linq.XDocument, MyJournal.Notebook.Config.IPageSettingsModel>;

namespace MyJournal.Notebook.Templates
{
    /// <summary>
    /// Base implementation for the OneNote page template presentation logic.
    /// </summary>
    abstract class PageTemplate : IPageTemplatePresenter, ITemplateMethod
    {
        protected OneNote.IApplication Application { get; private set; }
        protected string BulletSymbol { get; private set; }
        protected TaskScheduler ExclusiveScheduler { get; private set; }
        protected TaskFactory ExclusiveTaskFactory { get; private set; }

        /// <summary>
        /// OneNote XML Namespace.
        /// </summary>
        protected XNamespace OneNS { get; private set; }

        internal PageTemplate(OneNote.IApplication application)
        {
            try
            {
                Application = application;

                var i = 0;
                var symbol =
                    Component.AppSettings[TEMPLATES_BULLET_SYMBOL_PROPERTY];
                BulletSymbol = int.TryParse(symbol, out i) ? symbol : "2";

                ExclusiveScheduler =
                    new ConcurrentExclusiveSchedulerPair().ExclusiveScheduler;

                ExclusiveTaskFactory = new TaskFactory(ExclusiveScheduler);
                OneNS = application.GetXmlNamespace();
            }
            catch (Exception ex)
            {
                Utils.ExceptionHandler.HandleException(ex);
            }
        }

        #region IPageSettingsModel Events (Uses async ITemplateMethod impl)

        public void ChangeColor(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;
            FireAndForgetAsync(ChangeJournalPageColor, settings);
        }

        public void ChangeRuleLinesHorizontalColor(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;
            FireAndForgetAsync(ChangeJournalRuleLinesColor, settings);
        }

        public void ChangeRuleLinesHorizontalSpacing(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;
            FireAndForgetAsync(ChangeJournalRuleLinesSpacing, settings);
        }

        public void ChangeRuleLinesMarginColor(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;
            FireAndForgetAsync(ChangeJournalRuleLinesMarginColor, settings);
        }

        public void ChangeRuleLinesVisible(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;
            FireAndForgetAsync(ChangeJournalRuleLinesVisible, settings);
        }

        public void ChangeTitle(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;
            FireAndForgetAsync(ChangeJournalPageTitle, settings);
        }

        public async void CreateNewPage(object sender, EventArgs e)
        {
            var settings = sender as IPageSettingsModel;
            if (settings == null) return;

            int retryCount = 0, retryLimit = 4;
            Func<OneNote.Window, string[], Task> modalDisplay = null;
            OneNote.Window modalOwner = null;
            string[] modalText = null;
            var delay = TimeSpan.FromMilliseconds(250);

            for (; ; ) // BEGIN Retry pattern
            {
                try
                {
                    await CreateNewPageAsync(settings).ConfigureAwait(false);
                    break;
                }
                catch (Exception ex)
                {
                    modalOwner = Application.Windows.CurrentWindow;
                    if (Utils.ExceptionHandler.IsTransientError(ex))
                    {
                        if (++retryCount > retryLimit)
                        {
                            Utils.ExceptionHandler.HandleException(ex);
                            modalDisplay =
                                Utils.WinHelper.DisplayWarningAsync;
                            modalText = new[] {
                                "OneNote is busy.",
                                "Please try again in a moment."
                            };
                            break;
                        }
                    }
                    else
                    {
                        Utils.ExceptionHandler.HandleException(ex);
                        modalDisplay = Utils.WinHelper.DisplayErrorAsync;
                        modalText = new[] {
                            "OneNote API Error",
                            Utils.ExceptionHandler.FormatHResult(ex.HResult)
                        };
                        break;
                    }
                }/* end catch */

                Tracer.WriteWarnLine("Transient error, sleeping {0} ms . . .",
                  delay.Milliseconds);

                await Task.Delay(delay).ConfigureAwait(false);
            }// *END* Retry pattern

            if (modalDisplay != null)
            {
                await modalDisplay(modalOwner, modalText).ConfigureAwait(false);
            }
        }

        async Task CreateNewPageAsync(IPageSettingsModel settings)
        {
            await Task.Run(() =>
            {
                var context = new PageContext(Application, settings.Title);
                if (context.PageNotFound())
                {
                    var sw = Stopwatch.StartNew();

                    AddJournalPageContent(context,
                        context.CreateNewPage(), settings);

                    context.NavigateToPage();
                    RunKeyboardMacro();

                    Tracer.WriteInfoLine("CreateNewPageAsync elapsed time: {0} ms",
                        sw.ElapsedMilliseconds);
                    //DEBUG*/ context.SaveCurrentPageToDisk();
                    //DEBUG*/ context.SendCurrentPageToClipboard();
                }
                else
                {
                    context.NavigateToPage();
                    ScrollToBottomOfPage();
                }
            }).ConfigureAwait(false);
        }
        #endregion

        /// <summary>
        /// Positions the input cursor at the end of the outline element.
        /// </summary>
        internal virtual void RunKeyboardMacro()
        {
            ScrollToBottomOfPage();
        }

        // When updating page content, this mechanism prevents overlapped I/O;
        // otherwise OneNote may throw a COMException:
        // hrLastModifiedDateDidNotMatch (HRESULT 0x80042010)
        // REF: https://msdn.microsoft.com/en-us/library/office/ff966472(v=office.14).aspx
        async void FireAndForgetAsync(TemplateMethod template,
                                      IPageSettingsModel settings)
        {
            Action action = () =>
            {
                try
                {
                    SynchronizationContext.Current?.OperationStarted();
                    var context = PageContext.CurrentWindow(Application);
                    XDocument page;
                    if (context.IsMyJournalNotebook(out page))
                    {
                        var sw = Stopwatch.StartNew();
                        template.Invoke(context, page, settings);
                        PageTemplate.ScrollToTopOfPage();
                        Tracer.WriteInfoLine("{0} elapsed time: {1} ms",
                            template.Method.Name, sw.ElapsedMilliseconds);
                    }
                }
                catch (Exception ex)
                {
                    Utils.ExceptionHandler.HandleException(ex);
                }
                finally
                {
                    SynchronizationContext.Current?.OperationCompleted();
                }
            };

            await ExclusiveTaskFactory.StartNew(action,
              CancellationToken.None,
              TaskCreationOptions.DenyChildAttach,
              ExclusiveScheduler
            ).ConfigureAwait(false);
        }

        protected static void ScrollToBottomOfPage() =>
          System.Windows.Forms.SendKeys.SendWait(ShortcutKey.CTRL_END);

        protected static void ScrollToTopOfPage() =>
          System.Windows.Forms.SendKeys.SendWait(ShortcutKey.CTRL_HOME);

        #region Template Methods

        /// <summary>
        /// The template method for adding content to a journal page.
        /// </summary>
        public abstract void AddJournalPageContent(PageContext context,
          XDocument page, IPageSettingsModel settings);

        /// <summary>
        /// The template method for changing the journal page color.
        /// </summary>
        public abstract void ChangeJournalPageColor(PageContext context,
          XDocument page, IPageSettingsModel settings);

        /// <summary>
        /// The template method for changing the journal page rule lines color.
        /// </summary>
        public abstract void ChangeJournalRuleLinesColor(PageContext context,
          XDocument page, IPageSettingsModel settings);

        /// <summary>
        /// The template method for changing the journal page rule lines spacing.
        /// </summary>
        public abstract void ChangeJournalRuleLinesSpacing(PageContext context,
          XDocument page, IPageSettingsModel settings);

        /// <summary>
        /// The template method for changing the journal page rule lines margin color.
        /// </summary>
        public abstract void ChangeJournalRuleLinesMarginColor(PageContext context,
          XDocument page, IPageSettingsModel settings);

        /// <summary>
        /// The template method for changing the journal page rule lines visibility.
        /// </summary>
        public abstract void ChangeJournalRuleLinesVisible(PageContext context,
          XDocument page, IPageSettingsModel settings);

        /// <summary>
        /// The template method for changing the journal page title.
        /// </summary>
        public abstract void ChangeJournalPageTitle(PageContext context,
          XDocument page, IPageSettingsModel settings);

        #endregion

        internal const string
          TEMPLATES_BULLET_SYMBOL_PROPERTY = "Templates.Bullet.Symbol";
    }
}
