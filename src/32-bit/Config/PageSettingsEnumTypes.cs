using System;
using System.Xml.Serialization;

namespace MyJournal.Notebook.Config
{
    internal static class EnumExtensions
    {
        static readonly Type s_xmlEnumType = typeof(XmlEnumAttribute);

        internal static string XmlEnumValue(this Enum e)
        {
            var info = e.GetType().GetField(e.ToString("G"));
            var value = string.Empty;

            if (info.IsDefined(s_xmlEnumType, false))
            {
                var obj = info.GetCustomAttributes(s_xmlEnumType, false);
                var attrib = obj[0] as XmlEnumAttribute;
                value = attrib.Name;
            }

            return value;
        }
    }

    public enum PageColorEnum
    {
        [XmlEnum("#EDF5FE")]
        Blue,
        [XmlEnum("#FFFADF")]
        Yellow,
        [XmlEnum("#ECF0E5")]
        Green,
        [XmlEnum("#FFEEEF")]
        Red,
        [XmlEnum("#F3E9FF")]
        Purple,
        [XmlEnum("#E5EFEB")]
        Cyan,
        [XmlEnum("#FBF2E1")]
        Orange,
        [XmlEnum("#FBEDF6")]
        Magenta,
        [XmlEnum("#E5EDF3")]
        Blue_Mist,
        [XmlEnum("#E8E3EB")]
        Purple_Mist,
        [XmlEnum("#F9F4E7")]
        Tan,
        [XmlEnum("#FDFDDD")]
        Lemon,
        [XmlEnum("#ECF9E5")]
        Apple,
        [XmlEnum("#D4F9F2")]
        Teal,
        [XmlEnum("#EFDEDE")]
        Red_Chalk,
        [XmlEnum("#E9E9ED")]
        Silver,
        [XmlEnum("automatic")]
        None
    }

    public enum PageTitleEnum
    {
        [XmlEnum("dd-MM-yyyy")]
        DashDelimitedDate_DD_MM_YYYY,
        [XmlEnum("MM-dd-yyyy")]
        DashDelimitedDate_MM_DD_YYYY,
        [XmlEnum("yyyy-MM-dd")]
        DashDelimitedDate_YYYY_MM_DD,
        [XmlEnum("dd/MM/yyyy")]
        SlashDelimitedDate_DD_MM_YYYY,
        [XmlEnum("MM/dd/yyyy")]
        SlashDelimitedDate_MM_DD_YYYY,
        [XmlEnum("yyyy/MM/dd")]
        SlashDelimitedDate_YYYY_MM_DD,
        [XmlEnum("ddd d")]
        DayOfMonthDate_DDD_DD,
        [XmlEnum("dddd d")]
        DayOfMonthDate_DDDD_DD,
        [XmlEnum("MMMM d")]
        DayOfMonthDate_MMMM_DD
    }

    public enum RuleLinesColorEnum
    {
        [XmlEnum("#CAEBFD")]
        Light_Blue,
    }

    public enum RuleLinesMarginColorEnum
    {
        [XmlEnum("#FF5050")]
        Red,
    }

    public enum RuleLinesSpacingEnum
    {
        [XmlEnum("13.42771530151367")]
        Narrow_Ruled,
        [XmlEnum("23.76000022888184")]
        College_Ruled,
        [XmlEnum("33.11999893188476")]
        Standard_Ruled,
        [XmlEnum("46.79999923706054")]
        Wide_Ruled,
    }
}
