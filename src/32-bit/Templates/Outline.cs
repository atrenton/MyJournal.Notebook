using System;
using System.Xml.Linq;

namespace MyJournal.Notebook.Templates
{
    /// <summary>
    /// Sets the position and size of the OneNote Outline element, based on the
    /// user paper size preference.
    /// </summary>
    static class Outline
    {
        /// <summary>
        /// Positions the Outline note container on the page.
        /// All dimensions are specified in points.
        /// </summary>
        /// <param name="outline">Outline XML element</param>
        internal static void SetPosition(XElement outline)
        {
            var paperSize = Properties.Settings.Default.PaperSize;
            string xPosition, yPosition;

            switch (paperSize)
            {
                case PageSize.Automatic:
                    xPosition = "36.0";
                    yPosition = "68.11086273193359";
                    break;
                case PageSize.Letter:
                case PageSize.Legal:
                case PageSize.A4:
                    xPosition = "72.0";
                    yPosition = "89.71086883544922";
                    break;
                default:
                    throw new NotSupportedException(
                        $"Unknown paper size: {paperSize}");
            }

            var one = outline.Name.Namespace;
            outline.Add(
              new XElement(one + "Position",
                new XAttribute("x", xPosition),
                new XAttribute("y", yPosition),
                new XAttribute("z", "0")
              )
            );
        }

        /// <summary>
        /// Sets the Outline note container size.
        /// Assumes the default OneNote Calibri 11pt font is used.
        /// width = (paper width - 2" margins) * 72 points per inch
        /// height = 11pt * 1.5 line height
        /// </summary>
        /// <param name="outline">Outline XML element</param>
        internal static void SetSize(XElement outline)
        {
            var paperSize = Properties.Settings.Default.PaperSize;
            string width;

            switch (paperSize)
            {
                case PageSize.Automatic:
                case PageSize.Letter:
                case PageSize.Legal:
                    width = "468";
                    break;
                case PageSize.A4:
                    width = "451";
                    break;
                default:
                    throw new NotSupportedException(
                        $"Unknown paper size: {paperSize}");
            }

            var one = outline.Name.Namespace;
            outline.Add(
              new XElement(one + "Size",
                new XAttribute("width", width),
                new XAttribute("height", "16.5"),
                new XAttribute("isSetByUser", "true")
              )
            );
        }
    }
}
