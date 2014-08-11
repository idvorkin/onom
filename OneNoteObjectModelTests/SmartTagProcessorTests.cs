using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;
using NUnit.Framework;
using OnenoteCapabilities;
using OneNoteObjectModel;
using List = NUnit.Framework.List;

namespace OneNoteObjectModelTests
{
    // tests cases for our poor man XML parsing. 
    [TestFixture]
    public class SmartTagXMLProcessingTests
    {
        public SmartTagXMLProcessingTests()
        {
            _cursor = new OneNotePageCursor();
        }

        private OneNotePageCursor _cursor;

        [Test] 
        public void TestNotTag()
        {
            var unprocessedTag = "notATag Tag2";
            Assert.That(!SmartTag.IsSmartTag(unprocessedTag));
        }
    }

    [TestFixture]
    public class SmartTagProcessorTests
    {
        private TemporaryNoteBookHelper smartTagNoteBook;
        private OneNoteApp ona;

        [SetUp]
        public void Setup()
        {
            this.ona = new OneNoteApp();
            this._templateNotebook = new TemporaryNoteBookHelper(ona, "SmartTagTemplates");
            smartTagNoteBook = new TemporaryNoteBookHelper(ona, "SmartTag");
            this._settingsSmartTags = new SettingsSmartTags()
            {
                TemplateNotebook = _templateNotebook.Get().name,
                SmartTagStorageNotebook = smartTagNoteBook.Get().name
            };

            // create template structure.
            var templateSection = ona.CreateSection(_templateNotebook.Get(), _settingsSmartTags.TemplateSection);
            // create template page.
            _templateNotebook.CreatePage(new SmartTagModelTemplateContent(), _settingsSmartTags.SmartTagTemplateName,templateSection);

            // create smartTag model structure.
            ona.CreateSection(smartTagNoteBook.Get(), _settingsSmartTags.SmartTagModelSection);

            // Create a scratch page write on.
            this.pageContent = smartTagNoteBook.CreatePage(new SmartTagTestsPageConent(), "ContentPage");
            this.page = OneNoteApp.XMLDeserialize<Page>(pageContent.ToString());

            _cursor = new OneNotePageCursor();
            this.smartTagAugmenter = new SmartTagAugmenter(ona, _settingsSmartTags, new List<ISmartTagProcessor>());
        }


        // TODO: Come up with a better template for such tests.

        private SmartTagAugmenter smartTagAugmenter;
        private XDocument pageContent;
        private TemporaryNoteBookHelper _templateNotebook;
        private SettingsSmartTags _settingsSmartTags;
        private OneNotePageCursor _cursor;
        private Page page;

        [Test]
        public void EnumerateSmartTags()
        {
            var smartTags = SmartTag.Get(pageContent, _cursor).ToArray();
            Assert.AreEqual(2, smartTags.Count());
            Assert.AreEqual(1, smartTags.Count(st => st.IsProcessed()));
        }

        [Test]
        public void TestCompleting()
        {
            var smartTags = SmartTag.Get(pageContent, _cursor);
            foreach (var st in smartTags)
            {
                st.SetCompleted(ona);
            }
        }

        [Test]
        public void TestAugmentPage()
        {
            var smartTags = SmartTag.Get(this.pageContent, _cursor).ToArray();
            Assert.AreEqual(2,smartTags.Count());
            Assert.AreEqual(1, smartTags.Count(st => st.IsProcessed()));

            this.smartTagAugmenter.AugmentPage(ona, pageContent, new OneNotePageCursor());

            var smartTagsPostAugement = SmartTag.Get(this.pageContent, _cursor).ToArray();
            Assert.AreEqual(2, smartTagsPostAugement.Count());
            Assert.That(smartTagsPostAugement.All(st => st.IsProcessed()));
        }

        [Test]
        public void ProcessSmartTagAndLinkToOneNotePage()
        {
            var smartTag = SmartTag.Get(this.pageContent, _cursor).First(st => !st.IsProcessed());
            Assert.That(!smartTag.IsProcessed());

            // Brittle test, manually executing augmentation.
            smartTagAugmenter.AddToModel(smartTag, this.pageContent);

            // Set link to a random page.
            smartTag.SetLinkToPageId(ona, page.ID);

            Assert.That(smartTag.IsProcessed());
        }
        [Test]
        public void ProcessSmartTagAndLinkToURI()
        {
            var smartTag = SmartTag.Get(this.pageContent, _cursor).First(st => !st.IsProcessed());
            Assert.That(!smartTag.IsProcessed());
            smartTagAugmenter.AddToModel(smartTag, this.pageContent);
            Assert.That(smartTag.IsProcessed());

            // Set link to URI.
            smartTag.SetLink(ona, new Uri("http://helloworld"));
        }


        [TearDown]
        public void TearDown()
        {
            smartTagNoteBook.Dispose();
            _templateNotebook.Dispose();
        }
    }
}
