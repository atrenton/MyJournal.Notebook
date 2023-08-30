using System;
using System.Drawing;
using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Properties;
using Svg;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Templates
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance",
        "CA1812:Avoid uninstantiated internal classes")]
    class RetroPageTemplate : OrdinalDayPageTemplate
    {
        static readonly Font s_baseFont;
        static readonly Font s_superFont;

        string[] PageTitle { get; set; }

#pragma warning disable CA1810
        static RetroPageTemplate()
        {
            const FontStyle Style = FontStyle.Italic;
            var emSize = float.Parse(FONT_SIZE);
            s_baseFont = new(FONT_FAMILY_NAME, emSize, Style, PT);
            s_superFont = new(FONT_FAMILY_NAME, emSize / 2, Style, PT);
        }
#pragma warning restore CA1810

        internal RetroPageTemplate(OneNote.IApplication application)
          : base(application) { }

        /// <summary>
        /// Adds background image to page title.
        /// </summary>
        public override void AddJournalPageContent(PageContext context,
          XDocument page, IPageSettingsModel settings)
        {
            base.AddJournalPageContent(context, page, settings);

            AddTitleImage(page);
        }

        private void AddTitleImage(XDocument page)
        {
            Tracer.WriteTraceMethodLine();

            var svg = Asset.LoadSvgDocument("retro-title-background.svg");
            var viewport = MeasureTitleViewport(svg);

            page.Root.Element(OneNS + "Title").AddAfterSelf(
                new XElement(OneNS + "Image",
                    new XAttribute("format", "png"),
                    new XAttribute("backgroundImage", "true"),
                    new XElement(OneNS + "Position",
                        new XAttribute("x", viewport.X),
                        new XAttribute("y", viewport.Y)
                    ),
                    new XElement(OneNS + "Size",
                        new XAttribute("width", viewport.Width),
                        new XAttribute("height", viewport.Height),
                        new XAttribute("isSetByUser", "true")
                    ),
                    new XElement(OneNS + "Data", Asset.GetBase64Data(svg))
                )
            );
            Asset.DumpSvgElementAsync(svg);
        }

        /// <summary>
        /// Gets the page title background image viewport size.
        /// </summary>
        /// <param name="svg">SVG document reference</param>
        private Rectangle MeasureTitleViewport(SvgDocument svg)
        {
            var viewport = new Rectangle(23, 10, 223, 48); // in points

            using var g = Graphics.FromHwnd(IntPtr.Zero);
            g.PageUnit = PX;

            var pageTitleWidth = MeasureTitleWidth(g);

            var thrWidth = pageTitleWidth + THR_PADDING_RIGHT;

            if (thrWidth > THR_DEFAULT_WIDTH)
            {
                // Calculate SVG document values based on page title width.
                // SVG rectangle dimensions are at 2x scale to support 200%
                // zoom of raster image.

                var dxWidth = thrWidth - THR_DEFAULT_WIDTH; // @ 100% zoom

                var rect = svg.Children.GetSvgElementOf<SvgRectangle>();
                var rx = rect.CornerRadiusX.Value;

                var dx = (int)MathF.Ceiling(dxWidth / (rx / 2));

                svg.Width = new SvgUnit(PIXELS, svg.Width.Value + (dxWidth * 2));
                rect.Width = new SvgUnit(PIXELS, rect.Width.Value + (dxWidth * 2));

                // Resize viewport width (in points)
                viewport.Width += (int)MathF.Ceiling(dxWidth * 72 / g.DpiX) + dx;
            }

            return viewport;
        }

        private float MeasureTitleWidth(Graphics g)
        {
            var width = 0.0F;

            if (null != PageTitle)
            {
                var baseWidth = g.MeasureString(PageTitle[0], s_baseFont).Width;

                var superWidth = (PageTitle.Length > 1) ?
                    g.MeasureString(PageTitle[1], s_superFont).Width : 0.0F;

                width = MathF.Ceiling(baseWidth + superWidth);
            }

            return width;
        }

        protected override void SetPageTitle(XDocument page, string title, string style)
        {
            base.SetPageTitle(page, title, style);
            PageTitle = SplitTitle(title);
        }

        // Couldn't have done this without PowerToys Screen Ruler!
        // REF: https://learn.microsoft.com/en-us/windows/powertoys/screen-ruler

        /// <summary>
        /// Page Title Horizontal Rule (THR) default width, in pixels.
        /// </summary>
        const float
            THR_DEFAULT_WIDTH = 291.0F;   // 288px + 3px right padding

        /// <summary>
        /// Page Title Horizontal Rule (THR) padding on right, in pixels.
        /// </summary>
        const float THR_PADDING_RIGHT = 55.0F;

        const GraphicsUnit
            PT = GraphicsUnit.Point,
            PX = GraphicsUnit.Pixel;

        const SvgUnitType PIXELS = SvgUnitType.Pixel;
    }
}
