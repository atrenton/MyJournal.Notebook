using System;
using System.IO;
using System.Reflection;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Diagnostics;
using Office = Microsoft.Office.Core;

namespace MyJournal.Notebook.UI
{
    /// <summary>
    /// Implements a front controller for the Ribbon UI.
    /// </summary>
    class RibbonController : IRibbonEventHandler
    {
        const string DELEGATE_TO_VIEW = "delegate to RibbonView method = {0}";

        #region Properties

        internal RibbonView RibbonView { get; private set; }

        public Office.IRibbonUI RibbonUI { get; private set; }

        #endregion

        internal RibbonController()
        {
            CheckBoxClick = HandleCheckBoxAction;
            GetCheckBoxPressed = HandleGetPressed;
            GetToggleButtonPressed = HandleGetPressed;
            GetScreenTip = HandleGetScreenTip;
            RibbonLoad = ribbonUI => RibbonUI = ribbonUI;
            Startup = HandleStartupEvent;
            Shutdown = HandleShutdownEvent;
        }

        #region Controller Delegates

        void HandleCheckBoxAction(object sender, RibbonEventArgs e)
        {
            var methodName = e.Control.Tag + "_Click";
            Tracer.WriteTraceMethodLine(DELEGATE_TO_VIEW, methodName);
            var flags = BindingFlags.Instance | BindingFlags.NonPublic;
            var mi = RibbonView.GetType().GetMethod(methodName, flags);
            if (mi != null)
            {
                var checkBoxAction = (Action<object, RibbonEventArgs>)
                  Delegate.CreateDelegate(
                    typeof(Action<object, RibbonEventArgs>), RibbonView, mi);

                checkBoxAction(sender, e);

                switch (e.Control.Tag)
                {
                    case "PageTitle":
                        RibbonView.PageTitleControls.ForEach
                          (item => RibbonUI.InvalidateControl(item.ToString()));
                        break;
                    case "RuleLinesSpacing":
                        RibbonView.RuleLinesSpacingControls.ForEach
                          (item => RibbonUI.InvalidateControl(item.ToString()));
                        RibbonUI.InvalidateControl("No_Ruled_Lines");
                        break;
                }
            }
            else
            {
                Tracer.WriteErrorLine("Method '{0}' not found", methodName);
            }
        }

        bool HandleGetPressed(Office.IRibbonControl control)
        {
            var methodName = control.Tag + "_GetPressed";
            Tracer.WriteTraceMethodLine(DELEGATE_TO_VIEW, methodName);
            return DelegateToView<bool>(control, methodName);
        }

        string HandleGetScreenTip(Office.IRibbonControl control)
        {
            var methodName = control.Tag + "_GetScreenTip";
            Tracer.WriteTraceMethodLine(DELEGATE_TO_VIEW, methodName);
            return DelegateToView<string>(control, methodName);
        }

        void HandleStartupEvent(object sender, EventArgs e)
        {
            Tracer.WriteTraceMethodLine();
            if (sender is API.IOneNoteAddIn)
            {
                var application = ((API.IOneNoteAddIn)sender).Application;
                RibbonView = new RibbonView(application);

                ButtonClick += RibbonView.NotebookButton_Click;
                ButtonClick += RibbonView.PageColor_None;
                ColorClick += RibbonView.PageColor_Click;
                GetSuperTip += RibbonView.GetSuperTip;
                LoadImage += RibbonView.PageColor_LoadImage;
                ShowForm += RibbonView.OptionsForm_Show;
                ToggleButtonClick += (snd, evt) =>
                {
                    RibbonView.RuleLinesVisible_Toggle(snd, evt);

                    RibbonView.RuleLinesSpacingControls.ForEach(
                        item => RibbonUI.InvalidateControl(item.ToString()));
                };

                RibbonUI?.ActivateTabMso("TabHome");
            }
        }

        void HandleShutdownEvent(object sender, EventArgs e)
        {
            Tracer.WriteTraceMethodLine();
            if (RibbonView != null)
            {
                RibbonView.Dispose();
                RibbonView = null;
            }
        }

        #endregion

        TResult DelegateToView<TResult>(
          Office.IRibbonControl control, string methodName)
        {
            var result = default(TResult);
            var flags = BindingFlags.Instance | BindingFlags.NonPublic;
            var mi = RibbonView.GetType().GetMethod(methodName, flags);

            if (mi != null)
            {
                var viewFunc = (Func<Office.IRibbonControl, TResult>)
                  Delegate.CreateDelegate(
                    typeof(Func<Office.IRibbonControl, TResult>), RibbonView, mi);

                return viewFunc(control);
            }
            Tracer.WriteErrorLine("Method '{0}' not found", methodName);
            return result;
        }

        #region IRibbonEventHandler Members

        public void OnButtonClick(object sender, RibbonEventArgs e) =>
            ButtonClick?.Invoke(sender, e);

        public void OnCheckBoxClick(object sender, RibbonEventArgs e) =>
            CheckBoxClick?.Invoke(sender, e);

        public void OnColorClick(object sender, RibbonEventArgs e) =>
            ColorClick?.Invoke(sender, e);

        public bool OnCheckBoxGetPressed(Office.IRibbonControl control) =>
            GetCheckBoxPressed(control);

        public bool OnToggleButtonGetPressed(Office.IRibbonControl control) =>
            GetToggleButtonPressed(control);

        public System.IO.Stream OnLoadImage(string image) =>
            LoadImage?.Invoke(image);

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI) =>
            RibbonLoad?.Invoke(ribbonUI);

        public void OnShowForm(object sender, RibbonEventArgs e) =>
            ShowForm?.Invoke(sender, e);

        public void OnStartup(object sender, EventArgs e) =>
            Startup?.Invoke(sender, e);

        public void OnShutdown(object sender, EventArgs e) =>
            Shutdown?.Invoke(sender, e);

        public string OnScreentip(Office.IRibbonControl control) =>
            GetScreenTip?.Invoke(control);

        public string OnSupertip(Office.IRibbonControl control) =>
            GetSuperTip?.Invoke(control);

        public void OnToggleButtonClick(object sender, RibbonEventArgs e) =>
            ToggleButtonClick?.Invoke(sender, e);

        public event EventHandler<RibbonEventArgs> ButtonClick;

        public event EventHandler<RibbonEventArgs> CheckBoxClick;

        public event EventHandler<RibbonEventArgs> ColorClick;

        public event Func<Office.IRibbonControl, bool> GetCheckBoxPressed;

        public event Func<Office.IRibbonControl, string> GetScreenTip;

        public event Func<Office.IRibbonControl, string> GetSuperTip;

        public event Func<Office.IRibbonControl, bool> GetToggleButtonPressed;

        public event Func<string, MemoryStream> LoadImage;

        public event Action<Microsoft.Office.Core.IRibbonUI> RibbonLoad;

        public event EventHandler<RibbonEventArgs> ShowForm;

        public event EventHandler Startup;

        public event EventHandler Shutdown;

        public event EventHandler<RibbonEventArgs> ToggleButtonClick;

        #endregion
    }
}
