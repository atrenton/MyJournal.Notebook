using System;
using System.Xml.Linq;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Diagnostics;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.Templates
{
    class HappyDayPageTemplate : OrdinalDayPageTemplate
    {
        internal HappyDayPageTemplate(OneNote.IApplication application)
          : base(application) { }

        /// <summary>
        /// Creates Happy Day page content.
        /// </summary>
        /// <returns>OneNote page Outline XML element</returns>
        protected override XElement CreatePageContent()
        {
            Tracer.WriteTraceMethodLine();
            var day = DateTime.Today.ToString("dddd");
            var outline = base.CreateOutlineElement($"[Happy {day}]");

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

        /// <summary>
        /// Selects page content for replacement.
        /// </summary>
        internal override void RunKeyboardMacro()
        {
            const string Keys = ShortcutKey.CTRL_END + ShortcutKey.SHIFT_HOME;
            System.Windows.Forms.SendKeys.SendWait(Keys);
        }
    }
}
