using System;
using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Utils;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Templates
{
    class OrdinalDayPageTemplate : DefaultPageTemplate
    {
        // https://learn.microsoft.com/en-us/typography/font-list/lucida-handwriting
        // https://bigelowandholmes.typepad.com/bigelow-holmes/2014/10/how-and-why-we-designed-lucida.html
        protected const string
            FONT_FAMILY_NAME = "Lucida Handwriting",
            FONT_SIZE = "20";

        internal new const string PAGE_TITLE_FONT_STYLE =
            $"font-family:'{FONT_FAMILY_NAME}';font-size:{FONT_SIZE}.0pt";

        internal OrdinalDayPageTemplate(OneNote.IApplication application)
          : base(application) { }

        public override void ChangeJournalPageTitle(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();
            var created = (DateTime)page.Root.Attribute("dateTime");
            var title = created.Format(settings.Title);

            SetPageTitle(page, title, PAGE_TITLE_FONT_STYLE);
            context.UpdateMyJournal(page);

            UpdateRuledLinesView(context, settings);
        }

        protected override void SetPageTitle(XDocument page, string title, string style)
        {
            Tracer.WriteTraceMethodLine("title = {0}", title);
            if (title != null)
            {
                var styledTitle = StylizeTitle(title);

                var fontStyle = (styledTitle == title) ?
                    DefaultPageTemplate.PAGE_TITLE_FONT_STYLE :
                    OrdinalDayPageTemplate.PAGE_TITLE_FONT_STYLE;

                base.SetPageTitle(page, styledTitle, fontStyle);
            }
        }

        static string StylizeTitle(string title)
        {
            var parts = SplitTitle(title);
            if (parts.Length == 1) return parts[0]; // skip CSS styling

            var titleParts = $"title = {parts[0]}, superscript={parts[1]}";
            Tracer.WriteTraceMethodLine(titleParts);

            return $"{parts[0]}<span style='vertical-align:super'>{parts[1]}</span>";
        }

        static protected string[] SplitTitle(string title)
        {
            var length = char.IsDigit(title[0]) ? 1 : 2;
            var result = new string[length];
            if (length == 1)
            {
                result[0] = title;
            }
            else
            {
                result[0] = title[..^2];
                result[1] = title[^2..];
            }
            return result;
        }
    }
}
