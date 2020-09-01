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

        /// <summary>
        /// Select the whole line (when the cursor is at the end of the line).
        /// </summary>
        internal const string SHIFT_HOME = "+{HOME}";
    }
}
