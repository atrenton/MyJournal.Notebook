﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;

using App = MyJournal.Notebook.API.ApplicationExtensions;
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

            int retryCount = 0, retryLimit = 5;
            Func<OneNote.Window, string[], Task> modalDisplay = null;
            OneNote.Window modalOwner = null;
            string[] modalText = null;

            var latencyFactor = (StorageAccount.IsDefault) ? 1 : 10;
            var delay = TimeSpan.FromMilliseconds(500 * latencyFactor);
            var stopwatch = Stopwatch.StartNew();

            for (; ; ) // BEGIN Retry pattern
            {
                try
                {
                    await CreateNewPageTask(settings).ConfigureAwait(false);
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
                        var errorMessage = new List<string>
                        {
                            "OneNote API Error",
                            Utils.ExceptionHandler.FormatHResult(ex.HResult)
                        };
                        if (App.IsErrorCode(ex.HResult))
                        {
                            errorMessage.Add(App.ErrorCodeTable[ex.HResult]);
                        }
                        modalDisplay = Utils.WinHelper.DisplayErrorAsync;
                        modalText = errorMessage.ToArray();
                        break;
                    }
                }/* end catch */

                Tracer.WriteWarnLine("Transient error, sleeping {0} ms . . .",
                  delay.TotalMilliseconds);

                await Task.Delay(delay).ConfigureAwait(false);
            }// *END* Retry pattern

            stopwatch.Stop();

            if (modalDisplay != null)
            {
                await modalDisplay(modalOwner, modalText).ConfigureAwait(false);
            }
            else
            {
                Tracer.WriteInfoLine("CreateNewPage elapsed time: {0} ms",
                    stopwatch.ElapsedMilliseconds.ToString("N0"));
            }
        }

        Task CreateNewPageTask(IPageSettingsModel settings)
        {
            Action action = () =>
            {
                var context = new PageContext(Application, settings.Title);
                if (context.PageNotFound())
                {
                    var page = context.CreateNewPage();
                    AddJournalPageContent(context, page, settings);

                    // DateTime.MinValue is used to tell OneNote to skip its optimistic
                    // concurrency check; this is the initial page update.
                    context.UpdateMyJournal(page, DateTime.MinValue);

                    context.NavigateToPage();

                    UpdateRuledLinesView(context, settings);
                    RunKeyboardMacro();

                    //DEBUG*/ context.SaveCurrentPageToDisk();
                    //DEBUG*/ context.SendCurrentPageToClipboard();
                }
                else
                {
                    context.NavigateToPage();
                    context.SetFocus();
                    ScrollToBottomOfPage();
                }
            };

            var tcs = new TaskCompletionSource<object>();
            var thread = new Thread(() =>
                {
                    try
                    {
                        action();
                        tcs.SetResult(null);
                    }
                    catch (Exception e)
                    {
                        tcs.SetException(e);
                    }
                });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }
        #endregion

        // When updating page content, this mechanism prevents overlapped I/O;
        // otherwise OneNote may throw a COMException:
        // hrLastModifiedDateDidNotMatch (HRESULT 0x80042010)
        // REF: https://learn.microsoft.com/en-us/office/client-developer/onenote/error-codes-onenote
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

                        var one = page.Root.Name.Namespace;
                        var outline = page.Root.Element(one + "Outline");
                        if (Outline.IsNotEmpty(outline))
                        {
                            PageTemplate.ScrollToTopOfPage();
                        }

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

        /// <summary>
        /// Positions the input cursor at the end of the outline element.
        /// </summary>
        internal virtual void RunKeyboardMacro()
        {
            ScrollToBottomOfPage();
        }

        #region Keystroke Methods

        protected static void ScrollToBottomOfPage() =>
          System.Windows.Forms.SendKeys.SendWait(ShortcutKey.CTRL_END);

        protected static void ScrollToTopOfPage() =>
          System.Windows.Forms.SendKeys.SendWait(ShortcutKey.CTRL_HOME);

        // This method is a bit of a kludge.  It lets OneNote update the page
        // with a 1" clear top margin instead of rule lines all the way up.
        protected static void UpdateRuledLinesView(PageContext context,
            IPageSettingsModel settings)
        {
            if (settings.RuleLinesVisible)
            {
                var keys = settings.RuleLinesHorizontalSpacing switch
                {
                    RuleLinesSpacingEnum.Narrow_Ruled =>
                        ShortcutKey.SELECT_NARROW_RULED_LINES,

                    RuleLinesSpacingEnum.College_Ruled =>
                        ShortcutKey.SELECT_COLLEGE_RULED_LINES,

                    RuleLinesSpacingEnum.Standard_Ruled =>
                        ShortcutKey.SELECT_STANDARD_RULED_LINES,

                    RuleLinesSpacingEnum.Wide_Ruled =>
                        ShortcutKey.SELECT_WIDE_RULED_LINES,

                    _ => ShortcutKey.SELECT_NO_RULED_LINES
                };
                keys += ShortcutKey.SELECT_HOME_TAB;

                context.SetFocus();
                System.Windows.Forms.SendKeys.SendWait(keys);
            }
        }

        #endregion

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
