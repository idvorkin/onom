using System;
using System.Collections.Generic;
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
        private TemporaryNoteBookHelper _tempNoteBookHelperHelper;
        private Notebook tempNotebook;
            
        [TestFixtureSetUp]
        public void Setup()
        {
            _tempNoteBookHelperHelper = new TemporaryNoteBookHelper("CRUD");
            tempNotebook = _tempNoteBookHelperHelper.Get();
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            _tempNoteBookHelperHelper.Dispose();
        }

        [Test]
        public void CreateAndDelete()
        {
            // This fixture starts by creating and delete a notebook, so this NO-OP function makes sure create and delete works.
            return;
        }

        [Test]
        public void IsSamePage()
        {

            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNoteApplication.Instance.CreateSection(tempNotebook, sectionName);
            var page1 = OneNoteApplication.Instance.GetPageContentAsXDocument(OneNoteApplication.Instance.CreatePage(newSection, "Page1"));
            var page2 = OneNoteApplication.Instance.GetPageContentAsXDocument(OneNoteApplication.Instance.CreatePage(newSection, "Page2"));
            Assert.That(OneNoteApplication.IsSamePage(page1, page1));
            Assert.That(!OneNoteApplication.IsSamePage(page1, page2));

            var page1Again =
                OneNoteApplication.Instance.GetPageContentAsXDocument(
                    tempNotebook.PopulatedSection(sectionName).Page.First(p => p.name == "Page1"));
            Assert.That(OneNoteApplication.IsSamePage(page1, page1Again));
        }

        [Test]
        public void SectionDotGetPageTests()
        {
            var sectionName = Guid.NewGuid().ToString() + "GetPageTest";
            var newSection = OneNoteApplication.Instance.CreateSection(tempNotebook, sectionName);

            // should throw because page isn't present.
            Assert.Throws<InvalidOperationException>(() => newSection.GetPage("nonExistantPage"));

            var createdPage = OneNoteApplication.Instance.CreatePage(newSection, "newPage");
            var gottenPage = newSection.GetPage(pageName: createdPage.name);
            Assert.That(createdPage.name, Is.EqualTo(gottenPage.name));
        }

        [Test]
        public void ListNoteBooks()
        {
            var notebooks = OneNoteApplication.Instance.GetNotebooks();
            Assert.That(notebooks.Notebook.Any(), "You should have notebooks in Onenote");
            Assert.That(notebooks.Notebook.Any(n=>n.name == tempNotebook.name), "You should have the notebook we created");
        }

        [Test]
        public void AddSection()
        {
            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNoteApplication.Instance.CreateSection(tempNotebook, sectionName);
            Assert.That(newSection.name, Is.EqualTo(sectionName));
            Assert.That(OneNoteApplication.Instance.GetSections(tempNotebook).Any(s => s.name == sectionName));
        }

        [Test]
        public void AddPage()
        {
            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNoteApplication.Instance.CreateSection(tempNotebook, sectionName);

            var newPageName = Guid.NewGuid().ToString();
            var newPage1 = OneNoteApplication.Instance.CreatePage(newSection, newPageName);
            Assert.That( OneNoteApplication.Instance.GetSections(tempNotebook).First(s => s.name == sectionName).Page.Any(p => p.ID == newPage1.ID && p.name == newPageName));
        }

        [Test]
        public void ClonePage()
        {
            var sectionName = Guid.NewGuid().ToString();
            var newSection = OneNoteApplication.Instance.CreateSection(tempNotebook, sectionName);

            var firstPageTime = DateTime.Now - TimeSpan.FromHours(1.0);
            var newPage1 = OneNoteApplication.Instance.CreatePage(newSection, Guid.NewGuid().ToString());
            Assert.That( OneNoteApplication.Instance.GetSections(tempNotebook).First(s => s.name == sectionName).Page.Any(p => p.ID == newPage1.ID));
            newPage1.dateTime = firstPageTime;
            OneNoteApplication.Instance.UpdatePage(newPage1);

            var newPage2 = OneNoteApplication.Instance.ClonePage(newSection,newPage1,"NewTitle");
            Assert.That( OneNoteApplication.Instance.GetSections(tempNotebook).First(s => s.name == sectionName).Page.Any(p => p.ID == newPage2.ID), "New ID not set");

            Assert.That( OneNoteApplication.Instance.GetSections(tempNotebook).First(s => s.name == sectionName).Page.Any(p => p.name == "NewTitle"), "Title Not Update");

            Assert.That( OneNoteApplication.Instance.GetSections(tempNotebook).First(s => s.name == sectionName).Page.First(p => p.name == "NewTitle").dateTime != firstPageTime, "Page Creation Time Not Updated ");
        }

    }
}
