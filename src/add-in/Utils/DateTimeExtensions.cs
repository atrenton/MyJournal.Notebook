using System;
using MyJournal.Notebook.Config;

namespace MyJournal.Notebook.Utils
{
    static class DateTimeExtensions
    {
        /// <summary>
        /// Formats a DateTime value using a format specified by the PageTitleEnum
        /// XmlEnumAttribute decorator value.
        /// </summary>
        internal static string Format(this DateTime dt, PageTitleEnum pageTitle)
        {
            var dtLocal = dt.ToLocalTime();
            var result = dtLocal.ToString(pageTitle.XmlEnumValue());
            switch (pageTitle)
            {
                case PageTitleEnum.DayOfMonthDate_DDD_DD:
                case PageTitleEnum.DayOfMonthDate_DDDD_DD:
                case PageTitleEnum.DayOfMonthDate_MMMM_DD:
                    return string.Concat(result, dtLocal.OrdinalDaySuffix());
                default:
                    return result;
            }
        }

        internal static string OrdinalDaySuffix(this DateTime dt)
        {
            var day = dt.Day;
            var suffix = string.Empty;

            if (day % 31 >= 11 && day % 31 <= 13)
            {
                suffix = "th";
            }
            else
            {
                switch (day % 10)
                {
                    case 1: suffix = "st"; break;
                    case 2: suffix = "nd"; break;
                    case 3: suffix = "rd"; break;
                    default: suffix = "th"; break;
                }
            }
            return suffix;
        }
    }
}
