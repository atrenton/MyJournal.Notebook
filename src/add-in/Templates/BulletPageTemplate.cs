using System.Xml.Linq;
using MyJournal.Notebook.Diagnostics;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Templates
{
    class BulletPageTemplate : DefaultPageTemplate
    {
        internal BulletPageTemplate(OneNote.IApplication application)
          : base(application) { }

        /// <summary>
        /// Creates Bullet point page content.
        /// </summary>
        /// <returns>OneNote page Outline XML element</returns>
        protected override XElement CreatePageContent()
        {
            Tracer.WriteTraceMethodLine();
            var outline = base.CreatePageContent();

            // Insert a bullet symbol in the outline element (OE)
            var oe = outline.Element(OneNS + "OEChildren").Element(OneNS + "OE");
            oe.AddFirst(
              new XElement(OneNS + "List",
                  new XElement(OneNS + "Bullet",
                      new XAttribute("bullet", base.BulletSymbol),
                      new XAttribute("fontSize", "11")
                  )
              )
            );
            return outline;
        }
    }
}
