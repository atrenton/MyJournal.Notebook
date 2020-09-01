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
        internal new const string
        PAGE_TITLE_FONT_STYLE = "font-family:'Lucida Handwriting';font-size:20.0pt";

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

        static string[] SplitTitle(string title)
        {
            var length = char.IsDigit(title[0]) ? 1 : 2;
            var result = new string[length];
            if (length == 1)
            {
                result[0] = title;
            }
            else
            {
                var offset = title.Length - 2;
                result[0] = title.Substring(0, offset);
                result[1] = title.Substring(offset);
            }
            return result;
        }
    }
}
