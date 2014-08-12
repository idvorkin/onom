using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    public interface IPageContentAsText
    {
        string firstLine();
        string restOfLines();
    }

    public class TemporaryNoteBookHelper:IDisposable
    {
        private Notebook notebook;
        public Section DefaultSection;
        private string noteBookName;
        private string noteBookDirectory;

        public TemporaryNoteBookHelper(string TestName)
        {
            noteBookName = string.Format("onomTest_{0}_{1}", TestName, Guid.NewGuid());
            noteBookDirectory =  Path.Combine(Path.GetTempPath(),noteBookName);
            Directory.CreateDirectory(this.noteBookDirectory);
            notebook = OneNoteApplication.Instance.CreateNoteBook(noteBookDirectory, noteBookName);
            DefaultSection = OneNoteApplication.Instance.CreateSection(notebook, "defaultSection");
        }
        public XDocument CreatePage(IPageContentAsText pageContentAsText, string pageName, Section section=null)
        {
            if (section == null)
            {
                section = DefaultSection;
            }

            var p = OneNoteApplication.Instance.CreatePage(section, pageName);
            // the copied page has a PageID, replace it with PageID from the newly created page.
            var filledInPageHeader = string.Format(pageContentAsText.firstLine(), p.ID, pageName);
            var pageContent = XDocument.Parse(filledInPageHeader + pageContentAsText.restOfLines());
            OneNoteApplication.Instance.InteropApplication.UpdatePageContent(pageContent.ToString());
            // load the page back in as ObjectId's will be set.
            return OneNoteApplication.Instance.GetPageContentAsXDocument(p);
        }

        public Notebook Get()
        {
            return notebook;
        }
        public void Dispose()
        {
            OneNoteApplication.Instance.InteropApplication.CloseNotebook(OneNoteApplication.Instance.GetNotebooks().Notebook.First(n=>n.name == noteBookName).ID);
            Directory.Delete(noteBookDirectory,recursive:true);
        }
    }
}
