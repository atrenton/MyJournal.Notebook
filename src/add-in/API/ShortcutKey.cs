namespace MyJournal.Notebook.API
{
    /// <summary>
    /// A subset of keyboard shortcuts recognized by OneNote for Windows.
    /// </summary>
    static class ShortcutKey
    {
        /// <summary>Scroll to the bottom of the current page.</summary>
        internal const string CTRL_END = "^{END}";

        /// <summary>Scroll to the top of the current page.</summary>
        internal const string CTRL_HOME = "^{HOME}";

        /// <summary>Selects the Home tab on the ribbon.</summary>
        internal const string SELECT_HOME_TAB = "%h{ESC 2}";

        /// <summary>
        /// Selects No Ruled lines from the View tab.</summary>
        internal const string SELECT_NO_RULED_LINES = "%(wre)";

        /// <summary>Selects Narrow Ruled lines from the View tab.</summary>
        internal const string SELECT_NARROW_RULED_LINES = "%(wrn)";

        /// <summary>Selects College Ruled lines from the View tab.</summary>
        internal const string SELECT_COLLEGE_RULED_LINES = "%(wrc)";

        /// <summary>Selects Standard Ruled lines from the View tab.</summary>
        internal const string SELECT_STANDARD_RULED_LINES = "%(wrs)";

        /// <summary>Selects Wide Ruled lines from the View tab.</summary>
        internal const string SELECT_WIDE_RULED_LINES = "%(wrw)";

        /// <summary>
        /// Select the whole line (when the cursor is at the end of the line).
        /// </summary>
        internal const string SHIFT_HOME = "+{HOME}";
    }
}
