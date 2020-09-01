using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;

namespace MyJournal.Notebook.Templates
{
    /// <summary>
    /// Page template presentation method commands.
    /// </summary>
    interface ITemplateMethod
    {
        void AddJournalPageContent(PageContext context, XDocument page,
          IPageSettingsModel settings);

        void ChangeJournalPageColor(PageContext context, XDocument page,
          IPageSettingsModel settings);

        void ChangeJournalRuleLinesColor(PageContext context, XDocument page,
          IPageSettingsModel settings);

        void ChangeJournalRuleLinesSpacing(PageContext context, XDocument page,
          IPageSettingsModel settings);

        void ChangeJournalRuleLinesMarginColor(PageContext context, XDocument page,
          IPageSettingsModel settings);

        void ChangeJournalRuleLinesVisible(PageContext context, XDocument page,
          IPageSettingsModel settings);

        void ChangeJournalPageTitle(PageContext context, XDocument page,
          IPageSettingsModel settings);
    }
}
