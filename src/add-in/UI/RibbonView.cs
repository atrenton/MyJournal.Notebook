using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Resources;
using MyJournal.Notebook.API;
using MyJournal.Notebook.Config;
using MyJournal.Notebook.Diagnostics;
using MyJournal.Notebook.Properties;
using MyJournal.Notebook.Utils;
using Office = Microsoft.Office.Core;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.UI
{
    /// <summary>
    /// Implements the view for the Ribbon UI.
    /// </summary>
    class RibbonView : IDisposable
    {
        Binder _binder;
        bool _disposed;

        internal RibbonView(OneNote.IApplication application)
        {
            Tracer.WriteTraceTypeLine();
            System.Windows.Forms.Application.EnableVisualStyles();
            _binder = new Binder(application);
        }

        ~RibbonView() => Dispose();

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
                _binder.Dispose();
                _binder = null;
            }
            _disposed = true;
        }

        #endregion

        #region Properties

        internal static readonly
        List<PageColorEnum> PageColorControls =
          EnumHelper.ListOfEnumValues<PageColorEnum>();

        internal static readonly
        List<PageTitleEnum> PageTitleControls =
          EnumHelper.ListOfEnumValues<PageTitleEnum>();

        internal static readonly
        List<RuleLinesColorEnum> RuleLinesColorControls =
          EnumHelper.ListOfEnumValues<RuleLinesColorEnum>();

        internal static readonly
        List<RuleLinesSpacingEnum> RuleLinesSpacingControls =
          EnumHelper.ListOfEnumValues<RuleLinesSpacingEnum>();

        internal static readonly
        List<RuleLinesMarginColorEnum> RuleLinesMarginColorControls =
          EnumHelper.ListOfEnumValues<RuleLinesMarginColorEnum>();

        internal static readonly ResourceManager Resource =
          Resources.ResourceManager;

        #endregion

        #region Ribbon Event Handlers

        internal static string GetSuperTip(Office.IRibbonControl control)
        {
            Tracer.WriteTraceMethodLine($"Id = {control.Id}");
            var resourceName = $"UI_{control.Id}_Supertip";
            return Resource.GetString(resourceName, Resources.Culture);
        }

        internal void OptionsForm_Show(object sender, RibbonEventArgs e)
        {
            if (sender is API.AddInBase && e.Control.Tag == "OptionsForm")
            {
                Tracer.WriteTraceMethodLine();
                var addIn = sender as API.AddInBase;
                var savedTemplate = Properties.Settings.Default.PageTemplate;
                using (var form = new UI.OptionsForm())
                {
                    form.AboutIconImage_ResourceName = Component.AppIcon_ResourceName;
                    form.AboutText = string.Join(Environment.NewLine, new[] {
                          addIn.Description,
                          Component.AssemblyInfo.Copyright,
                          $"Version {Component.AssemblyInfo.ProductVersion}"
                        });
                    form.Title = $"{Component.AssemblyInfo.Title} Options";

                    var owner = e.Control.Context as OneNote.Window;
                    form.ShowDialog(new Utils.WinHelper.Win32Window(owner));
                }
                if (Properties.Settings.Default.PageTemplate != savedTemplate)
                {
                    _binder.Rebind(addIn.Application);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        internal void NotebookButton_Click(object sender, RibbonEventArgs e)
        {
            if (e.Control.Id == "NotebookButton")
            {
                _binder.CreateJournalPage();
            }
        }

        internal void PageColor_Click(object sender, RibbonEventArgs e)
        {
            Tracer.WriteTraceMethodLine();

            if (e.Control.Id == "PageColorGallery")
            {
                var fmt = "Id = {0}, selectedId = {1}, selectedIndex = {2}";
                Tracer.WriteDebugLine(fmt, e.Control.Id, e.SelectedId, e.SelectedIndex);
                _binder.PageSettings.Color = PageColorControls[e.SelectedIndex];
            }
        }

        internal static MemoryStream PageColor_LoadImage(string imageName)
        {
            Tracer.WriteTraceMethodLine();

            var stream = new MemoryStream();
            using (var image = GetSolidColorImage(imageName))
            {
                image.Save(stream, ImageFormat.Bmp);
            }
            return stream;
        }

        internal void PageColor_None(object sender, RibbonEventArgs e)
        {
            if (e.Control.Id == "PageColorNone")
            {
                Tracer.WriteTraceMethodLine();
                _binder.PageSettings.Color = PageColorEnum.None;
            }
        }

        internal void PageTitle_Click(object sender, RibbonEventArgs e)
        {
            if (e.Control.Tag != "PageTitle") return;

            Tracer.WriteTraceMethodLine();
            var id = e.Control.Id;
            Tracer.WriteDebugLine("Id = {0}, Pressed = {1}", id, e.Pressed);

            if (e.Pressed)
            {
                var t = typeof(PageTitleEnum);
                if (Enum.IsDefined(t, id))
                {
                    _binder.PageSettings.Title = EnumHelper.Parse<PageTitleEnum>(id);
                }
                else
                {
                    Tracer.WriteErrorLine("'{0}' is not a valid {1} value", id, t.Name);
                }
            }
            else // set default value
            {
                _binder.PageSettings.Title = PageTitleEnum.SlashDelimitedDate_MM_DD_YYYY;
            }
        }

        internal bool PageTitle_GetPressed(Office.IRibbonControl control)
        {
            Tracer.WriteTraceMethodLine();
            var pressed = false;
            if (control.Tag == "PageTitle")
            {
                var pageTitleEnum = EnumHelper.Parse<PageTitleEnum>(control.Id);
                pressed = (_binder.PageSettings.Title == pageTitleEnum);
            }
            return pressed;
        }

#pragma warning disable CA1822
        internal string PageTitle_GetScreenTip(Office.IRibbonControl control)
#pragma warning restore CA1822
        {
            Tracer.WriteTraceMethodLine("Id = {0}", control.Id);
            var result = string.Empty;
            if (control.Tag == "PageTitle")
            {
                var pageTitleEnum = EnumHelper.Parse<PageTitleEnum>(control.Id);
                result = DateTime.Now.Format(pageTitleEnum);
            }
            return result;
        }

        internal void RuleLinesSpacing_Click(object sender, RibbonEventArgs e)
        {
            if (e.Control.Tag != "RuleLinesSpacing") return;

            Tracer.WriteTraceMethodLine();
            var id = e.Control.Id;
            Tracer.WriteDebugLine("Id = {0}, Pressed = {1}", id, e.Pressed);

            if (e.Pressed)
            {
                var t = typeof(RuleLinesSpacingEnum);
                if (Enum.IsDefined(t, id))
                {
                    var ruleLinesSpacingEnum = EnumHelper.Parse<RuleLinesSpacingEnum>(id);
                    _binder.PageSettings.RuleLinesHorizontalSpacing = ruleLinesSpacingEnum;
                    _binder.PageSettings.RuleLinesVisible = true;
                }
                else
                {
                    Tracer.WriteErrorLine("'{0}' is not a valid {1} value", id, t.Name);
                }
            }
            else // set default value
            {
                _binder.PageSettings.RuleLinesVisible = false;
            }
        }

        internal bool RuleLinesSpacing_GetPressed(Office.IRibbonControl control)
        {
            Tracer.WriteTraceMethodLine();
            var pressed = false;

            if (control.Tag == "RuleLinesSpacing"
            && _binder.PageSettings.RuleLinesVisible)
            {
                var ruleLinesSpacingEnum = EnumHelper.Parse<RuleLinesSpacingEnum>(control.Id);
                pressed =
                (_binder.PageSettings.RuleLinesHorizontalSpacing == ruleLinesSpacingEnum);
            }
            return pressed;
        }

        internal bool RuleLinesVisible_GetPressed(Office.IRibbonControl control)
        {
            Tracer.WriteTraceMethodLine();
            var pressed = false;

            if (control.Tag == "RuleLinesVisible")
            {
                pressed = !_binder.PageSettings.RuleLinesVisible;
            }
            return pressed;
        }

        internal void RuleLinesVisible_Toggle(object sender, RibbonEventArgs e)
        {
            if (e.Control.Tag != "RuleLinesVisible") return;

            Tracer.WriteTraceMethodLine();
            var pressed = e.Pressed;

            _binder.PageSettings.RuleLinesVisible = !pressed;
            Tracer.WriteDebugLine("Id = {0}, Pressed = {1}", e.Control.Id, pressed);
        }

        #endregion

        private static Image GetSolidColorImage(string imageName)
        {
            try
            {
                var pageColor = EnumHelper.Parse<PageColorEnum>(imageName);
                var rgbValue = pageColor.XmlEnumValue();
                Tracer.WriteTraceMethodLine("Image name={0}, Color={1}",
                  imageName, rgbValue);

                const int X = 0, Y = 0, Width = 32, Height = 32;
                Image image = new Bitmap(Width, Height);
                var color = ColorTranslator.FromHtml(rgbValue);

                using (var g = Graphics.FromImage(image))
                {
                    using (Brush brush = new SolidBrush(color))
                    {
                        g.DrawRectangle(Pens.LightGray, X, Y, Width - 1, Height - 1);
                        g.FillRectangle(brush, X + 1, Y + 1, Width - 2, Height - 2);
                    }
                }

                return image;
            }
            catch (Exception e)
            {
                Utils.ExceptionHandler.HandleException(e);
            }
            return null;
        }
    }
}
