using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Xml;
using System.Xml.Serialization;
using MyJournal.Notebook.Diagnostics;

namespace MyJournal.Notebook.Config
{
    [XmlRoot(ElementName = "pageSettings", IsNullable = false,
    Namespace = "urn:MyJournal-Notebook-Config:PageSettings-v1.0")]
    public class PageSettings : IPageSettingsModel
    {
        const string
        PROP_CHG_MSG1 = "{0} {1} property changed:",
        PROP_CHG_MSG2 = "  OLD value: {0}",
        PROP_CHG_MSG3 = "  NEW value: {0}";

        #region Fields

        // Initialize default values for Add-in
        PageColorEnum _color = PageColorEnum.Apple;
        bool _ruleLinesVisible = true;
        RuleLinesColorEnum _ruleLinesHorizontalColor = RuleLinesColorEnum.Light_Blue;
        RuleLinesSpacingEnum _ruleLinesHorizontalSpacing = RuleLinesSpacingEnum.Narrow_Ruled;
        RuleLinesMarginColorEnum _ruleLinesMarginColor = RuleLinesMarginColorEnum.Red;
        PageTitleEnum _title = PageTitleEnum.DayOfMonthDate_MMMM_DD;

        bool _modified;

        #endregion

        #region IPageSettingsModel Events

        public event EventHandler ColorChanged;

        public event EventHandler RuleLinesHorizontalColorChanged;

        public event EventHandler RuleLinesHorizontalSpacingChanged;

        public event EventHandler RuleLinesMarginColorChanged;

        public event EventHandler RuleLinesVisibleChanged;

        public event EventHandler TitleChanged;

        #endregion

        #region Properties

        [XmlElement(ElementName = "color")]
        public PageColorEnum Color
        {
            get
            {
                return _color;
            }
            set
            {
                if (SetField(ref _color, value))
                {
                    OnColorChanged(EventArgs.Empty);
                }
            }
        }

        [XmlElement(DataType = "boolean", ElementName = "ruleLinesVisible")]
        public bool RuleLinesVisible
        {
            get
            {
                return _ruleLinesVisible;
            }
            set
            {
                if (SetField(ref _ruleLinesVisible, value))
                {
                    OnRuleLinesVisibleChanged(EventArgs.Empty);
                }
            }
        }

        [XmlElement(ElementName = "ruleLinesHorizontalColor")]
        public RuleLinesColorEnum RuleLinesHorizontalColor
        {
            get
            {
                return _ruleLinesHorizontalColor;
            }
            set
            {
                if (SetField(ref _ruleLinesHorizontalColor, value))
                {
                    OnRuleLinesHorizontalColorChanged(EventArgs.Empty);
                }
            }
        }

        [XmlElement(ElementName = "ruleLinesHorizontalSpacing")]
        public RuleLinesSpacingEnum RuleLinesHorizontalSpacing
        {
            get
            {
                return _ruleLinesHorizontalSpacing;
            }
            set
            {
                if (SetField(ref _ruleLinesHorizontalSpacing, value))
                {
                    OnRuleLinesHorizontalSpacingChanged(EventArgs.Empty);
                }
            }
        }

        [XmlElement(ElementName = "ruleLinesMarginColor")]
        public RuleLinesMarginColorEnum RuleLinesMarginColor
        {
            get
            {
                return _ruleLinesMarginColor;
            }
            set
            {
                if (SetField(ref _ruleLinesMarginColor, value))
                {
                    OnRuleLinesMarginColorChanged(EventArgs.Empty);
                }
            }
        }

        [XmlElement(ElementName = "title")]
        public PageTitleEnum Title
        {
            get
            {
                return _title;
            }
            set
            {
                if (SetField(ref _title, value))
                {
                    OnTitleChanged(EventArgs.Empty);
                }
            }
        }

        #endregion

        public PageSettings(bool setDefaultValues) : this()
        {
            _modified = setDefaultValues;
        }

        public PageSettings()
        {
            Tracer.WriteTraceTypeLine();
        }

        #region Methods

        void DisplayStateChange(object oldValue, object newValue,
          string propertyName)
        {
            var typeFullName = GetType().FullName;
            Tracer.WriteTraceLine(PROP_CHG_MSG1, typeFullName, propertyName);
            Tracer.WriteDittoLine(PROP_CHG_MSG2, oldValue);
            Tracer.WriteDittoLine(PROP_CHG_MSG3, newValue);
        }

        public bool IsModified() => _modified;

        protected virtual void OnColorChanged(EventArgs e) =>
            ColorChanged?.Invoke(this, e);

        protected virtual void OnRuleLinesHorizontalColorChanged(EventArgs e) =>
            RuleLinesHorizontalColorChanged?.Invoke(this, e);

        protected virtual void OnRuleLinesHorizontalSpacingChanged(EventArgs e) =>
            RuleLinesHorizontalSpacingChanged?.Invoke(this, e);

        protected virtual void OnRuleLinesMarginColorChanged(EventArgs e) =>
            RuleLinesMarginColorChanged?.Invoke(this, e);

        protected virtual void OnRuleLinesVisibleChanged(EventArgs e) =>
            RuleLinesVisibleChanged?.Invoke(this, e);

        protected virtual void OnTitleChanged(EventArgs e) =>
            TitleChanged?.Invoke(this, e);

        protected bool SetField<T>(ref T field, T value,
           [CallerMemberName] string propertyName = "")
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;

            // Check if object is being deserialized
            var callingMethod = new StackTrace().GetFrame(2).GetMethod();
            if (callingMethod.Module.Name == "<In Memory Module>")
            {
                field = value;
                return false;
            }

            DisplayStateChange(field, value, propertyName);
            field = value;
            return (_modified = true);
        }

        #endregion
    }
}
