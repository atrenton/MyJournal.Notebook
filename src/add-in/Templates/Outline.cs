using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;

namespace MyJournal.Notebook.Templates
{
    /// <summary>
    /// Inserts the cursor location, position and size for the OneNote Outline
    /// note container.
    /// </summary>
    static class Outline
    {
        /// <summary>
        /// Inserts the cursor location into an empty Outline note container.
        /// </summary>
        /// <param name="outline">Outline XML element</param>
        internal static void InsertCursor(XElement outline)
        {
            if ( IsEmpty(outline) && null != outline )
            {
                const string Attribute = "selected";
                var one = outline.Name.Namespace;
                var oeChildren = outline.Element(one + "OEChildren");
                var oe = oeChildren.Element(one + "OE");

                var elem = new List<XElement> {
                    outline.Parent, outline, oeChildren, oe
                };
                elem.ForEach(e => e.SetAttributeValue(Attribute, "partial"));

                oe.Element(one + "T").SetAttributeValue(Attribute, "all");
            }
        }

        /// <summary>
        /// Returns true if Outline note container has no content.
        /// </summary>
        /// <param name="outline">Outline XML element</param>
        internal static bool IsEmpty(XElement outline)
        {
            if (null == outline) return true;

            var one = outline.Name.Namespace;
            var oe = outline.Element(one + "OEChildren").Element(one + "OE");

            var firstNode = oe.FirstNode;
            if (firstNode.NodeType == XmlNodeType.Element)
            {
                var t = firstNode as XElement;
                if ("T" == t.Name.LocalName)
                {
                    var node = t.FirstNode;
                    if (node.NodeType == XmlNodeType.CDATA)
                    {
                        return string.IsNullOrEmpty(((XCData)node).Value);
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Returns true if Outline note container has content.
        /// </summary>
        /// <param name="outline">Outline XML element</param>
        internal static bool IsNotEmpty(XElement outline) => !IsEmpty(outline);

        /// <summary>
        /// Positions the Outline note container on the page.
        /// All dimensions are specified in points.
        /// Assumes the default OneNote Calibri 11pt font is used.
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
                    yPosition = "68.11085510253906";
                    break;
                case PageSize.Letter:
                case PageSize.Legal:
                case PageSize.A4:
                    xPosition = "72.0";
                    yPosition = "89.71085357666015";
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
