using System;
using System.Collections.Generic;

namespace MyJournal.Notebook.Utils
{
    /// <summary>
    /// Generic helper methods for Enum types
    /// </summary>
    static class EnumHelper
    {
        private static T[] GetEnumValues<T>() => (T[])Enum.GetValues(typeof(T));

        internal static List<T> ListOfEnumValues<T>() =>
            new List<T>(GetEnumValues<T>());

        internal static T Parse<T>(string id) => (T)Enum.Parse(typeof(T), id);
    }
}
