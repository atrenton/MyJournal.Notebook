using System;
using Office = Microsoft.Office.Core;

namespace MyJournal.Notebook.API
{
    public class RibbonEventArgs : EventArgs
    {
        public RibbonEventArgs(Office.IRibbonControl control)
          : this(control, false, string.Empty, -1)
        {
        }

        public RibbonEventArgs(Office.IRibbonControl control, bool pressed)
          : this(control, pressed, string.Empty, -1)
        {
        }

        public RibbonEventArgs(Office.IRibbonControl control,
          bool pressed, string selectedId, int selectedIndex) : base()
        {
            Control = control;
            Pressed = pressed;
            SelectedId = selectedId;
            SelectedIndex = selectedIndex;
        }

        #region Properties

        public Office.IRibbonControl Control
        {
            get;
            private set;
        }

        public bool Pressed
        {
            get;
            private set;
        }

        public string SelectedId
        {
            get;
            private set;
        }

        public int SelectedIndex
        {
            get;
            private set;
        }

        #endregion
    }
}
