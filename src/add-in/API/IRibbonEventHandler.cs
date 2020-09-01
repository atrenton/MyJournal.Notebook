using System;
using Office = Microsoft.Office.Core;

namespace MyJournal.Notebook.API
{
    public interface IRibbonEventHandler
    {
        event EventHandler<RibbonEventArgs> ButtonClick;
        event EventHandler<RibbonEventArgs> CheckBoxClick;
        event EventHandler<RibbonEventArgs> ColorClick;
        event Func<Office.IRibbonControl, bool> GetCheckBoxPressed;
        event Func<Office.IRibbonControl, string> GetScreenTip;
        event Func<Office.IRibbonControl, string> GetSuperTip;
        event Func<Office.IRibbonControl, bool> GetToggleButtonPressed;
        event Func<string, System.IO.MemoryStream> LoadImage;
        event Action<Office.IRibbonUI> RibbonLoad;
        event EventHandler<RibbonEventArgs> ShowForm;
        event EventHandler Startup;
        event EventHandler Shutdown;
        event EventHandler<RibbonEventArgs> ToggleButtonClick;

        void OnButtonClick(object sender, RibbonEventArgs e);
        void OnCheckBoxClick(object sender, RibbonEventArgs e);
        bool OnCheckBoxGetPressed(Office.IRibbonControl control);
        void OnColorClick(object sender, RibbonEventArgs e);
        System.IO.Stream OnLoadImage(string image);
        void OnRibbonLoad(Office.IRibbonUI ribbonUI);
        void OnShowForm(object sender, RibbonEventArgs e);
        void OnStartup(object sender, EventArgs e);
        void OnShutdown(object sender, EventArgs e);
        string OnScreentip(Office.IRibbonControl control);
        string OnSupertip(Office.IRibbonControl control);
        void OnToggleButtonClick(object sender, RibbonEventArgs e);
        bool OnToggleButtonGetPressed(Office.IRibbonControl control);
    }
}
