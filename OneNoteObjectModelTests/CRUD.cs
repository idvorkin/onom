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
using NUnit.Framework.Constraints;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    class CRUD
    {
        public OneNoteApp OneNote;

        // GRR - I don't understant why I can't open hierarchy, clearly I'm missing something.
        private static readonly string CurrentAssemblyPath = Path.GetDirectoryName(new Uri(Assembly.GetAssembly(typeof (OneNoteApp)).CodeBase).LocalPath);
        private static readonly string TestNoteBookPath = Path.Combine(CurrentAssemblyPath, @"..\..\..\TestNoteBooks\");
        private string tempTestNoteBookName;
        private string tempTestNoteBookDirectory;
        private Notebook testNotebook;
            
        [TestFixtureSetUp]
        public void Setup()
        {
            OneNote = new OneNoteApp();
            tempTestNoteBookName = "onomTest_" + Guid.NewGuid();
            tempTestNoteBookDirectory=  Path.Combine(Path.GetTempPath(),tempTestNoteBookName);
            Directory.CreateDirectory(tempTestNoteBookDirectory);
            testNotebook = OneNote.CreateNoteBook(tempTestNoteBookDirectory, tempTestNoteBookDirectory);
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
            var notebooks = OneNote.GetNotebooks();
            Assert.That(notebooks.Notebook.Any(), "You should have notebooks in Onenote");
            Assert.That(notebooks.Notebook.Any(n=>n.name == tempTestNoteBookName), "You should have the notebook we created");
        }

        [Test]
        public void AddSection()
        {
            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNote.CreateSection(testNotebook, sectionName);
            Assert.That(newSection.name, Is.EqualTo(sectionName));
            Assert.That(OneNote.GetSections(testNotebook).Any(s => s.name == sectionName));
        }

        [Test]
        public void AddPage()
        {
            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNote.CreateSection(testNotebook, sectionName);

            var newPage1 = OneNote.CreatePage(newSection);
            Assert.That( OneNote.GetSections(testNotebook).First(s => s.name == sectionName).Page.Any(p => p.ID == newPage1.ID));
        }

        [Test]
        public void ClonePage()
        {
            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNote.CreateSection(testNotebook, sectionName);

            var newPage1 = OneNote.CreatePage(newSection);
            Assert.That( OneNote.GetSections(testNotebook).First(s => s.name == sectionName).Page.Any(p => p.ID == newPage1.ID));

            var newPage2 = OneNote.ClonePage(newSection,newPage1);
            Assert.That( OneNote.GetSections(testNotebook).First(s => s.name == sectionName).Page.Any(p => p.ID == newPage2.ID));
        }

    }
}
