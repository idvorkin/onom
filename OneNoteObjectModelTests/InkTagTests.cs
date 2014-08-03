using System.Linq;
using System.Xml.Linq;
using NUnit.Framework;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class InkTagTests
    {
        private TemporaryNoteBookHelper inkTagNoteBook;
        private OneNoteApp ona;
        private XDocument pageContent;

        [SetUp]
        public void Setup()
        {
            this.ona = new OneNoteApp();
            inkTagNoteBook = new TemporaryNoteBookHelper(ona, "InkTag");
            this.pageContent = inkTagNoteBook.CreatePage(new InkTagsTestPageContent(), "InkPage");
        }

        [Test]
        public void ValidTestSetup()
        {
            // make sure page loads and parses.
            return;
        }

        [TearDown]
        public void TearDown()
        {
            inkTagNoteBook.Dispose();
        }
    }
}