using System;

namespace MyJournal.Notebook.API
{
    public interface ICOMAddIn
    {
        Microsoft.Office.Core.COMAddIn COMAddIn { get; }
    }
}
