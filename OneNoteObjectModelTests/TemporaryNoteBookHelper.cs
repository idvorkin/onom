using System;
using System.IO;
using System.Linq;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    class TemporaryNoteBookHelper:IDisposable
    {
        private OneNoteApp ona;
        private Notebook notebook;
        private string noteBookName;
        private string noteBookDirectory;

        public TemporaryNoteBookHelper(OneNoteApp ona)
        {
            this.ona = ona;
            noteBookName = "onomTest_" + Guid.NewGuid();
            noteBookDirectory =  Path.Combine(Path.GetTempPath(),noteBookName);
            Directory.CreateDirectory(this.noteBookDirectory);
            notebook = ona.CreateNoteBook(noteBookDirectory, noteBookName);
        }
        public Notebook Get()
        {
            return notebook;
        }
        public void Dispose()
        {
            ona.OneNoteApplication.CloseNotebook(ona.GetNotebooks().Notebook.First(n=>n.name == noteBookName).ID);
            Directory.Delete(noteBookDirectory,recursive:true);
        }
    }
}