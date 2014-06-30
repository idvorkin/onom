using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using NUnit.Framework;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class SmartTagProcessorTests
    {
        private TemporaryNoteBookHelper smartTagNoteBook;
        private OneNoteApp ona;

        [TestFixtureSetUp]
        public void Setup()
        {
            this.ona = new OneNoteApp();
            smartTagNoteBook = new TemporaryNoteBookHelper(ona);
            this._templateNotebook = new TemporaryNoteBookHelper(ona);
            this._settingsSmartTags = new SettingsSmartTags()
            {
                    TemplateNotebook = _templateNotebook.Get().name,
                    SmartTagNotebook =  smartTagNoteBook.Get().name
            };

            // create template structure.
            var templateSection  = ona.CreateSection(_templateNotebook.Get(), _settingsSmartTags.TemplateSection);
            ona.CreatePage(templateSection, _settingsSmartTags.SmartTagTemplateName);

            // create smartTag model structure.
            var modelSection  = ona.CreateSection(smartTagNoteBook.Get(), _settingsSmartTags.SmartTagStorageSection);

            this.defaultSection = ona.CreateSection(smartTagNoteBook.Get(), "default");
            var p = ona.CreatePage(defaultSection, Guid.NewGuid().ToString());

            // the copied page has a PageID, replace it with PageID from the newly created page.
            var filledInPage = string.Format(pageWithOneProcessedAndOneUnProcessedSmartTag, p.ID, p.name);
            this.pageContent = XDocument.Parse(filledInPage);
            ona.OneNoteApplication.UpdatePageContent(pageContent.ToString());

            this.smartTagAugmenter = new SmartTagAugmenter(ona, _settingsSmartTags, null);
        }

        // TODO: Come up with a better template for such tests.
        readonly string pageWithOneProcessedAndOneUnProcessedSmartTag = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"Scratch\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-06-28T19:00:46.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">"+
    "  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
    "  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
    "  <one:PageSettings RTL=\"false\" color=\"automatic\">"+
    "    <one:PageSize>"+
    "      <one:Automatic />"+
    "    </one:PageSize>"+
    "    <one:RuleLines visible=\"false\" />"+
    "  </one:PageSettings>"+
    "  <one:Title lang=\"en-US\">"+
    "    <one:OE author=\"Igor Dvorkin\" authorInitials=\"ID\" authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T19:00:35.000Z\" lastModifiedTime=\"2014-06-28T19:00:35.000Z\" alignment=\"left\" quickStyleIndex=\"0\">"+
    "      <one:T><![CDATA[{1}]]></one:T>"+
    "    </one:OE>"+
    "  </one:Title>"+
    "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-06-28T19:00:44.000Z\">"+
    "    <one:Position x=\"36.0\" y=\"86.4000015258789\" z=\"0\" />"+
    "    <one:Size width=\"167.2590484619141\" height=\"67.13858032226562\" />"+
    "    <one:OEChildren>"+
    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T19:00:35.000Z\" lastModifiedTime=\"2014-06-28T19:00:35.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
    "        <one:T><![CDATA[]]></one:T>"+
    "      </one:OE>"+
    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T19:00:35.000Z\" lastModifiedTime=\"2014-06-28T19:00:35.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
    "        <one:T><![CDATA[This is an empty page with a ]]></one:T>"+
    "      </one:OE>"+
    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T19:00:44.000Z\" lastModifiedTime=\"2014-06-28T19:00:44.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
    "        <one:T><![CDATA[#unprocessedTag Tag2]]></one:T>"+
    "      </one:OE>"+
    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T19:00:14.000Z\" lastModifiedTime=\"2014-06-28T19:00:35.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
    "        <one:T><![CDATA[<a"+
    "href=\"onenote:#processedTag|6/28/2014 12:00:31 PM&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch/SmartTagStorage.one\">#</a><a"+
    "href=\"onenote:#processedTag&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch/Current.one\">processedTag</a> tag1<a"+
    "href=\"http://ObjectId=e9692a18-92ef-47cc-b6e8-351626931e41\">.</a>]]></one:T>"+
    "      </one:OE>"+
    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T19:00:35.000Z\" lastModifiedTime=\"2014-06-28T19:00:35.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
    "        <one:T><![CDATA[]]></one:T>"+
    "      </one:OE>"+
    "    </one:OEChildren>"+
    "  </one:Outline>"+
    "</one:Page>";

        private SmartTagAugmenter smartTagAugmenter;
        private XDocument pageContent;
        private Section defaultSection;
        private TemporaryNoteBookHelper _templateNotebook;
        private SettingsSmartTags _settingsSmartTags;

        [Test]
        public void EnumerateSmartTags()
        {
            var smartTags = smartTagAugmenter.GetUnProcessedSmartTags(pageContent);
            Assert.That(smartTags.Count() == 1);
            Assert.That(smartTags.All(st => !st.IsProcessed()));
            return;
        }
        [Test]
        public void ProcessSmartTagAndLinkToOneNotePage()
        {
            var smartTag = smartTagAugmenter.GetUnProcessedSmartTags(this.pageContent).First();
            Assert.That(!smartTag.IsProcessed());
            smartTagAugmenter.AddToModel(smartTag,this.pageContent);
            Assert.That(smartTag.IsProcessed());

            smartTagAugmenter.AddLinkToSmartTag(smartTag, this.pageContent, this.defaultSection,"DeadPage"); 
            return;
        }
        [Test]
        [Ignore("Only one tag processor will pass tests - needs to be debugged.")]
        public void ProcessSmartTagAndLinkToURI()
        {
            var smartTag = smartTagAugmenter.GetUnProcessedSmartTags(this.pageContent).First();
            Assert.That(!smartTag.IsProcessed());
            smartTagAugmenter.AddToModel(smartTag,this.pageContent);
            Assert.That(smartTag.IsProcessed());

            // Set link to URI.
            smartTagAugmenter.AddLinkToSmartTag(smartTag,this.pageContent, new Uri("http://helloworld"));

            // Set link to Section Page.
            // smartTagAugmenter.AddLinkToSmartTag(smartTags.First(), this.pageContent, this.defaultSection,"DeadPage"); 
            return;
        }


        [TestFixtureTearDown]
        public void TearDown()
        {
            smartTagNoteBook.Dispose();
        }
    }
}
