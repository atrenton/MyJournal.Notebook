using System;

namespace MyJournal.Notebook.Config
{
    /// <summary>
    /// The page settings model implements the domain data properties for the
    /// journal page and provides one-way data binding.
    /// </summary>
    public interface IPageSettingsModel
    {
        event EventHandler ColorChanged;
        event EventHandler RuleLinesHorizontalColorChanged;
        event EventHandler RuleLinesHorizontalSpacingChanged;
        event EventHandler RuleLinesMarginColorChanged;
        event EventHandler RuleLinesVisibleChanged;
        event EventHandler TitleChanged;

        bool IsModified();

        PageColorEnum Color { get; set; }
        RuleLinesColorEnum RuleLinesHorizontalColor { get; set; }
        RuleLinesSpacingEnum RuleLinesHorizontalSpacing { get; set; }
        RuleLinesMarginColorEnum RuleLinesMarginColor { get; set; }
        bool RuleLinesVisible { get; set; }
        PageTitleEnum Title { get; set; }
    }
}
