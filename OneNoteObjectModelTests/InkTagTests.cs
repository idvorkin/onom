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
        private XDocument pageContent;

        [SetUp]
        public void Setup()
        {
            inkTagNoteBook = new TemporaryNoteBookHelper("InkTag");
            this.pageContent = inkTagNoteBook.CreatePage(new InkTagsTestPageContentInkCanada(), "InkPage");
        }

        [Test]
        public void TestPageWithNoInk()
        {
            var nonInkPage = inkTagNoteBook.CreatePage(new SmartTagTestsPageConent(), "PageWithNoInk");
            var inkTags = InkTag.Get(nonInkPage).ToArray();
            Assert.IsEmpty(inkTags);
        }

        [Test]
        public void TestEnumerate()
        {
            var inkTags = InkTag.Get(pageContent).ToArray();
            Assert.AreEqual(inkTags.Count(), 1);

            var inkTag = inkTags.First();
            Assert.AreEqual("#info Canada", inkTag.FullText);
            return;
        }

        [Test]
        public void TestTextReplacement()
        {
            var inkTags = InkTag.Get(pageContent).ToArray();
            var inkTag = inkTags.First();
            Assert.AreEqual("#info Canada", inkTag.FullText);

            inkTag.ToText();
            var newInkTags = InkTag.Get(pageContent).ToArray();
            Assert.IsEmpty(newInkTags);

            // should now have an info smartTag
            var smartTag = SmartTag.Get(pageContent, null).First();
            Assert.AreEqual("#info Canada",smartTag.FullText);
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
