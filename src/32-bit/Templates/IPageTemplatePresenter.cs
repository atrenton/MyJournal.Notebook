using System;

namespace MyJournal.Notebook.Templates
{
    /// <summary>
    /// Implements the page template presentation logic for the journal page.
    /// </summary>
    interface IPageTemplatePresenter
    {
        void ChangeColor(object sender, EventArgs e);
        void ChangeRuleLinesHorizontalColor(object sender, EventArgs e);
        void ChangeRuleLinesHorizontalSpacing(object sender, EventArgs e);
        void ChangeRuleLinesMarginColor(object sender, EventArgs e);
        void ChangeRuleLinesVisible(object sender, EventArgs e);
        void ChangeTitle(object sender, EventArgs e);
        void CreateNewPage(object sender, EventArgs e);
    }
}
