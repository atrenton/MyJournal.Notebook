using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Utils;
using Office = Microsoft.Office.Core;

namespace MyJournal.Notebook.API
{
    [ComVisible(true), Guid(ComClassId.AddInBase)]
    public abstract class AddInBase : ICOMAddIn, IOneNoteAddIn, IDisposable
    {
        bool _disposed;

        /// <summary>
        /// Specifies how this add-in is connected
        /// </summary>
        public ext_ConnectMode ConnectMode { get; protected set; }

        /// <summary>
        /// Handler for ribbon events fired by this add-in's callback methods
        /// </summary>
        public IRibbonEventHandler RibbonEventHandler { get; protected set; }

        protected AddInBase() { }

        ~AddInBase() => Dispose(false);

        #region ICOMAddIn Member

        public Microsoft.Office.Core.COMAddIn COMAddIn
        {
            get;
            protected set;
        }

        #endregion

        #region IOneNoteAddIn Member

        public Microsoft.Office.Interop.OneNote.IApplication Application
        {
            get;
            protected set;
        }

        #endregion

        #region IDisposable Member

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)  // dispose of managed resources
            {
                Tracer.WriteTraceMethodLine();
                Application = null;
                COMAddIn = null;
                RibbonEventHandler = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            _disposed = true;
        }

        #endregion

        public virtual string Description
        {
            get
            {
                return $"{COMAddIn?.Description} Add-In for OneNote";
            }
        }

        #region Callbacks

        public virtual string GetAboutLabel(Office.IRibbonControl control)
        {
            return COMAddIn.Description;
        }

        public virtual string GetAboutScreentip(Office.IRibbonControl control)
        {
            return string.Concat(COMAddIn.Description, " Add-In");
        }

        public virtual string GetOptionsScreentip(Office.IRibbonControl control)
        {
            return string.Concat(COMAddIn.Description, " Add-In Options");
        }

        public bool GetCheckBoxPressed(Office.IRibbonControl control)
        {
            return RibbonEventHandler.OnCheckBoxGetPressed(control);
        }

        public string GetScreentip(Office.IRibbonControl control)
        {
            return RibbonEventHandler.OnScreentip(control);
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            return RibbonEventHandler.OnSupertip(control);
        }

        public bool GetToggleButtonPressed(Office.IRibbonControl control)
        {
            return RibbonEventHandler.OnToggleButtonGetPressed(control);
        }

        public object LoadImage(string image) =>
            new ReadOnlyIStreamWrapper(RibbonEventHandler.OnLoadImage(image));

        public virtual void NotifyButtonAction(Office.IRibbonControl control)
        {
            RibbonEventHandler.OnButtonClick(this, new RibbonEventArgs(control));
        }

        public virtual void NotifyCheckBoxAction(Office.IRibbonControl control,
          bool pressed)
        {
            RibbonEventHandler.OnCheckBoxClick
              (this, new RibbonEventArgs(control, pressed));
        }

        public virtual void NotifyColorAction(Office.IRibbonControl control,
          string selectedId, int selectedIndex)
        {
            var evt = new RibbonEventArgs(control, false, selectedId, selectedIndex);
            RibbonEventHandler.OnColorClick(this, evt);
        }

        public void NotifyRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            Tracer.WriteTraceMethodLine();
            RibbonEventHandler.OnRibbonLoad(ribbonUI);
        }

        public virtual void NotifyShowForm(Office.IRibbonControl control)
        {
            RibbonEventHandler.OnShowForm(this, new RibbonEventArgs(control));
        }

        protected virtual void NotifyStartup()
        {
            Tracer.WriteTraceMethodLine();
            RibbonEventHandler.OnStartup(this, EventArgs.Empty);
        }

        protected virtual void NotifyShutdown()
        {
            Tracer.WriteTraceMethodLine();
            RibbonEventHandler.OnShutdown(this, EventArgs.Empty);
            Dispose();
        }

        public virtual void NotifyToggleButtonAction(Office.IRibbonControl control,
          bool pressed)
        {
            RibbonEventHandler.OnToggleButtonClick
              (this, new RibbonEventArgs(control, pressed));
        }

        #endregion
    }
}
