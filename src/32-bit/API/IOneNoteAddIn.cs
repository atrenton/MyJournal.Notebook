using OneNote = Microsoft.Office.Interop.OneNote;

namespace MyJournal.Notebook.API
{
    public interface IOneNoteAddIn
    {
        OneNote.IApplication Application { get; }
    }
}
