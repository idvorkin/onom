using System;
using System.IO;
using System.Linq;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    public class TemporaryNoteBookHelper:IDisposable
    {
        private OneNoteApp ona;
        private Notebook notebook;
        private string noteBookName;
        private string noteBookDirectory;

        public TemporaryNoteBookHelper(OneNoteApp ona, string TestName)
        {
            this.ona = ona;
            noteBookName = string.Format("onomTest_{0}_{1}", TestName, Guid.NewGuid());
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