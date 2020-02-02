using System;
using System.Xml.Linq;

namespace MyJournal.Notebook.Templates
{
    /// <summary>
    /// Sets the OneNote PageSize descendant elements, based on the user paper
    /// size preference.
    /// </summary>
    static class PageSize
    {
        internal const string
            Automatic = "Automatic",
            Letter = "Letter",
            Legal = "Legal",
            A4 = "A4";

        #region Properties

        /// <summary>
        /// Gets the collection of paper size names.
        /// </summary>
        internal static object[] Items =>
          new[] { Automatic, Letter, Legal, A4 };

        #endregion

        /// <summary>
        /// Sets the dimensions for the user selected paper size.
        /// Orientation defaults to portrait.
        /// Print margins default to 0.5" top & bottom; 1" left & right.
        /// </summary>
        /// <param name="pageSize">PageSize XML element</param>
        internal static void SetDimensions(XElement pageSize)
        {
            var paperSize = Properties.Settings.Default.PaperSize;
            Diagnostics.Tracer.WriteDebugLine("Paper size: {0}", paperSize);
            if (paperSize == Automatic) return;

            string width, height;
            switch (paperSize)
            {
                case Letter:
                    width = "612.0";
                    height = "792.0";
                    break;
                case Legal:
                    width = "612.0";
                    height = "1008.0";
                    break;
                case A4:
                    width = "595.2755737304687";
                    height = "841.8897094726562";
                    break;
                default:
                    throw new NotSupportedException(
                        $"Unknown paper size: {paperSize}");
            }

            var one = pageSize.Name.Namespace;
            pageSize.Descendants().Remove();
            pageSize.Add(
                new XElement(one + "Orientation",
                new XAttribute("landscape", "false")
                )
            );
            pageSize.Add(
                new XElement(one + "Dimensions",
                new XAttribute("width", width), new XAttribute("height", height)
                )
            );
            pageSize.Add(
                new XElement(one + "Margins",
                new XAttribute("top", "36.0"), new XAttribute("bottom", "36.0"),
                new XAttribute("left", "72.0"), new XAttribute("right", "72.0")
                )
            );
        }
    }
}
