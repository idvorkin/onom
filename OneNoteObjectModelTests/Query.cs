using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.OneNote;
using NUnit.Framework;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    class Query
    {
        public OneNoteApp OneNote;

        // GRR - I don't understant why I can't open hierarchy, clearly I'm missing something.
        private static readonly string CurrentAssemblyPath = Path.GetDirectoryName(new Uri(Assembly.GetAssembly(typeof (OneNoteApp)).CodeBase).LocalPath);
        private static readonly string TestNoteBookPath = Path.Combine(CurrentAssemblyPath, @"..\..\..\TestNoteBooks\");
        private string tempTestNoteBookName;
        private string tempTestNoteBookDirectory;
            
        [TestFixtureSetUp]
        public void Setup()
        {
            OneNote = new OneNoteApp();
            tempTestNoteBookName = "onomTest_" + Guid.NewGuid();
            tempTestNoteBookDirectory=  Path.Combine(Path.GetTempPath(),tempTestNoteBookName);
            Directory.CreateDirectory(tempTestNoteBookDirectory);
            OneNote.CreateNoteBook(tempTestNoteBookDirectory, tempTestNoteBookDirectory);
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            OneNote.OneNoteApplication.CloseNotebook(OneNote.GetNotebooks().Notebook.First(n=>n.name == tempTestNoteBookName).ID);
            Directory.Delete(tempTestNoteBookDirectory,recursive:true);
        }

        [Test]
        public void CreateAndDelete()
        {
            // This fixture starts by creating and delete a notebook, so this NO-OP function makes sure create and delete works.
            return;
        }

        [Test]
        public void ListNoteBooks()
        {
            /*
            var onePkgPath = Path.Combine(TestNoteBookPath, @"Simple\Section 1.one");
            Console.WriteLine(onePkgPath);
            string unusedNoteBookID=string.Empty;
            // OneNote.OneNoteApplication.CreateN
            OneNote.OneNoteApplication.OpenHierarchy(onePkgPath, "", out unusedNoteBookID, CreateFileType.cftNone);
            */
            var notebooks = OneNote.GetNotebooks();
            Assert.That(notebooks.Notebook.Any(), "You should have notebooks in Onenote");
            Assert.That(notebooks.Notebook.Any(n=>n.name == tempTestNoteBookName), "You should have the notebook we created");
            Console.WriteLine(TestNoteBookPath);
        }

    }
}
