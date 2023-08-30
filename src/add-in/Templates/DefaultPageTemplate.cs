using System;
using System.Linq;
using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Utils;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Templates
{
    class DefaultPageTemplate : PageTemplate
    {
        internal const string
          COLOR_ATTRIBUTE_NAME = "color",
          PAGE_TITLE_FONT_STYLE = "font-family:Calibri;font-size:18.0pt";

        internal DefaultPageTemplate(OneNote.IApplication application)
          : base(application) { }

        #region Template Methods

        public override void AddJournalPageContent(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();

            InitPageSettings(page, settings);

            SetPageTitle(page, context.PageName, PAGE_TITLE_FONT_STYLE);

            page.Root.Add(CreatePageContent());
            Outline.InsertCursor(page.Root.Element(OneNS + "Outline"));
        }

        public override void ChangeJournalPageColor(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();

            var element = page.Root.Element(OneNS + "PageSettings");
            var value = settings.Color.XmlEnumValue();
            element.SetAttributeValue(COLOR_ATTRIBUTE_NAME, value);

            context.UpdateMyJournal(page);

            UpdateRuledLinesView(context, settings);
        }

        public override void ChangeJournalRuleLinesColor(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();

            var element = page.Root.Element(OneNS + "PageSettings")
              .Element(OneNS + "RuleLines")
              .Element(OneNS + "Horizontal");
            var value = settings.RuleLinesHorizontalColor.XmlEnumValue();
            element.SetAttributeValue(COLOR_ATTRIBUTE_NAME, value);

            context.UpdateMyJournal(page);
        }

        public override void ChangeJournalRuleLinesSpacing(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();
            var pageSettings = page.Root.Element(OneNS + "PageSettings");
            var ruleLines = pageSettings.Element(OneNS + "RuleLines");

            if ((bool)ruleLines.Attribute("visible"))
            {
                var element = ruleLines.Element(OneNS + "Horizontal");
                var value = settings.RuleLinesHorizontalSpacing.XmlEnumValue();
                element.SetAttributeValue("spacing", value);
            }
            else
            {
                SetRuleLines(pageSettings, settings);
            }
            context.UpdateMyJournal(page);

            UpdateRuledLinesView(context, settings);
        }

        public override void ChangeJournalRuleLinesMarginColor(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();
            throw new NotImplementedException();
        }

        public override void ChangeJournalRuleLinesVisible(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            Tracer.WriteTraceMethodLine();
            SetRuleLines(page.Root.Element(OneNS + "PageSettings"), settings);
            context.UpdateMyJournal(page);

            UpdateRuledLinesView(context, settings);
        }

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

        #endregion

        #region Default Page Content Helpers

        /// <summary>
        /// Creates an Outline note container.
        /// </summary>
        /// <param name="content">User defined content</param>
        /// <returns>OneNote Outline XML element</returns>
        protected virtual XElement CreateOutlineElement(string content)
        {
            var outline = new XElement(OneNS + "Outline");
            Outline.SetPosition(outline);
            Outline.SetSize(outline);

            /* Add the child content to the Outline container */
            outline.Add(
              new XElement(OneNS + "OEChildren",
                new XElement(OneNS + "OE",
                  new XElement(OneNS + "T", new XCData(content))
                )
              )
            );
            return outline;
        }

        /// <summary>
        /// Creates empty page content.
        /// </summary>
        /// <returns>OneNote page Outline XML element</returns>
        protected virtual XElement CreatePageContent()
        {
            return CreateOutlineElement(string.Empty);
        }

        /// <summary>
        /// Sets the page size, rule lines style and background color.
        /// </summary>
        /// <param name="page">page content in XML format</param>
        /// <param name="pageModel">page settings data model</param>
        protected void InitPageSettings(XDocument page, IPageSettingsModel pageModel)
        {
            var pageSettings = page.Root.Element(OneNS + "PageSettings");
            if (pageSettings == null)
            {
                throw new ArgumentNullException(
                    nameof(page), "one:PageSettings element not found");
            }

            SetPageSize(pageSettings);
            SetRuleLines(pageSettings, pageModel);

            // Set Page background color
            pageSettings.SetAttributeValue(COLOR_ATTRIBUTE_NAME,
                                           pageModel.Color.XmlEnumValue());
        }

        /// <summary>
        /// Sets the page size to the user preferred paper size dimensions.
        /// </summary>
        /// <param name="pageSettings">PageSettings XML element</param>
        protected virtual void SetPageSize(XElement pageSettings)
        {
            var pageSize = pageSettings.Element(OneNS + "PageSize");
            if (pageSize == null)
            {
                throw new ArgumentNullException(
                    nameof(pageSettings), "one:PageSize element not found");
            }
            PageSize.SetDimensions(pageSize);
        }

        /// <summary>
        /// Sets the page title and its font style.
        /// </summary>
        /// <param name="page">page content in XML format</param>
        /// <param name="title">page title</param>
        /// <param name="style">font style</param>
        protected virtual void SetPageTitle(XDocument page, string title,
          string style)
        {
            var titleElement = page.Root.Element(OneNS + "Title");
            var t = titleElement?.Descendants(OneNS + "T").Last();
            if (t != null)
            {
                t.FirstNode.ReplaceWith(new XCData(title));
                if (!string.IsNullOrEmpty(style))
                {
                    t.Parent.SetAttributeValue("style", style);
                }
            }
        }

        /// <summary>
        /// Sets the page rule lines.
        /// </summary>
        /// <param name="pageSettings">PageSettings XML element</param>
        /// <param name="pageModel">page settings data model</param>
        protected void SetRuleLines(XElement pageSettings,
          IPageSettingsModel pageModel)
        {
            var ruleLines = pageSettings.Element(OneNS + "RuleLines");
            if (ruleLines == null)
            {
                throw new ArgumentNullException(
                    nameof(pageSettings), "one:RuleLines element not found");
            }
            ruleLines.Descendants().Remove();

            var visible = pageModel.RuleLinesVisible;
            if (visible)
            {
                ruleLines.Add(
                  new XElement(OneNS + "Horizontal",
                      new XAttribute(COLOR_ATTRIBUTE_NAME, pageModel
                                    .RuleLinesHorizontalColor.XmlEnumValue()),
                      new XAttribute("spacing", pageModel
                                    .RuleLinesHorizontalSpacing.XmlEnumValue())
                  )
                );
                ruleLines.Add(
                  new XElement(OneNS + "Margin",
                      new XAttribute(COLOR_ATTRIBUTE_NAME, pageModel
                                    .RuleLinesMarginColor.XmlEnumValue())
                  )
                );
            }
            ruleLines.SetAttributeValue("visible", visible);
        }

        #endregion
    }
}
