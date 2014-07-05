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
    // tests cases for our poor man XML parsing. 
    [TestFixture]
    public class SmartTagXMLProcessingTests
    {
        [Test] 
        public void TestNotTag()
        {
            var unprocessedTag = "notATag Tag2";
            Assert.That(!SmartTagAugmenter.IsSmartTag(unprocessedTag));
        }

        [Test] 
        public void TestSimpleTag()
        {
            var unprocessedTag = "#unprocessedTag Tag2";
            Assert.That(SmartTagAugmenter.IsSmartTag(unprocessedTag));

            var st = SmartTagAugmenter.SmartTagFromElement(unprocessedTag,null);
            Assert.That(st.TextAfterTag(), Is.EqualTo("Tag2"));
            Assert.That(st.TagName() , Is.EqualTo("unprocessedTag"));
        }

        [Test] 
        public void TestPunctuation()
        {
            var unprocessedTag = "#unprocessedTag Tag2 - ! , fad|";
            Assert.That(SmartTagAugmenter.IsSmartTag(unprocessedTag));

            var st = SmartTagAugmenter.SmartTagFromElement(unprocessedTag,null);
            Assert.That(st.TextAfterTag(), Is.EqualTo("Tag2 - ! , fad|"));
            Assert.That(st.TagName() , Is.EqualTo("unprocessedTag"));
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
            smartTagNoteBook = new TemporaryNoteBookHelper(ona, "SmartTag");
            this._templateNotebook = new TemporaryNoteBookHelper(ona, "SmartTagTemplates");
            this._settingsSmartTags = new SettingsSmartTags()
            {
                    TemplateNotebook = _templateNotebook.Get().name,
                    SmartTagStorageNotebook =  smartTagNoteBook.Get().name
            };

            // create template structure.
            var templateSection  = ona.CreateSection(_templateNotebook.Get(), _settingsSmartTags.TemplateSection);
            var smartTagTemplatePage = ona.CreatePage(templateSection, _settingsSmartTags.SmartTagTemplateName);


            // create smartTagModelTemplate
            // the copied page has a PageID, replace it with PageID from the newly created page.
            var filledInTempaltePage = string.Format(smartTagModelTemplatePageWithDumbTodoTable, smartTagTemplatePage.ID, smartTagTemplatePage.name);
            ona.OneNoteApplication.UpdatePageContent(filledInTempaltePage);

            // create smartTag model structure.
            ona.CreateSection(smartTagNoteBook.Get(), _settingsSmartTags.SmartTagStorageSection);

            this.defaultSection = ona.CreateSection(smartTagNoteBook.Get(), "section_to_find_content");
            var p = ona.CreatePage(defaultSection, Guid.NewGuid().ToString());

            // the copied page has a PageID, replace it with PageID from the newly created page.
            var filledInPage = string.Format(pageWithOneProcessedAndOneUnProcessedSmartTag, p.ID, p.name);
            this.pageContent = XDocument.Parse(filledInPage);
            ona.OneNoteApplication.UpdatePageContent(pageContent.ToString());

            this.smartTagAugmenter = new SmartTagAugmenter(ona, _settingsSmartTags, null);
        }

        readonly string smartTagModelTemplatePageWithDumbTodoTable = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">"+
"  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
"  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
"  <one:PageSettings RTL=\"false\" color=\"automatic\">"+
"    <one:PageSize>"+
"      <one:Automatic />"+
"    </one:PageSize>"+
"    <one:RuleLines visible=\"false\" />"+
"  </one:PageSettings>"+
"  <one:Title lang=\"en-US\">"+
"    <one:OE author=\"Igor Dvorkin\" authorInitials=\"ID\" authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-02T13:54:19.000Z\" lastModifiedTime=\"2014-07-02T13:54:19.000Z\" alignment=\"left\" quickStyleIndex=\"0\">"+
"      <one:T><![CDATA[{1}]]></one:T>"+
"    </one:OE>"+
"  </one:Title>"+
"  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\">"+
"    <one:Position x=\"54.0\" y=\"122.400001525879\" z=\"0\" />"+
"    <one:Size width=\"731.7882080078124\" height=\"35.37544250488281\" />"+
"    <one:OEChildren>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:14:07.000Z\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\" alignment=\"left\">"+
"        <one:Table bordersVisible=\"true\" hasHeaderRow=\"true\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\">"+
"          <one:Columns>"+
"            <one:Column index=\"0\" width=\"445.8600769042969\" isLocked=\"true\" />"+
"            <one:Column index=\"1\" width=\"277.7980651855469\" isLocked=\"true\" />"+
"          </one:Columns>"+
"          <one:Row lastModifiedTime=\"2014-07-05T03:15:19.000Z\">"+
"            <one:Cell lastModifiedTime=\"2014-07-05T03:15:19.000Z\" lastModifiedByInitials=\"ID\">"+
"              <one:OEChildren>"+
"                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:14:07.000Z\" lastModifiedTime=\"2014-07-05T03:15:19.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
"                  <one:T><![CDATA[<span"+
"style='font-weight:bold'>Task</span>]]></one:T>"+
"                </one:OE>"+
"              </one:OEChildren>"+
"            </one:Cell>"+
"            <one:Cell lastModifiedTime=\"2014-07-05T03:15:19.000Z\" lastModifiedByInitials=\"ID\">"+
"              <one:OEChildren>"+
"                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:14:07.000Z\" lastModifiedTime=\"2014-07-05T03:15:19.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
"                  <one:T><![CDATA[<span"+
"style='font-weight:bold'>Date</span>]]></one:T>"+
"                </one:OE>"+
"              </one:OEChildren>"+
"            </one:Cell>"+
"          </one:Row>"+
"        </one:Table>"+
"      </one:OE>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:15:27.000Z\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:T><![CDATA[]]></one:T>"+
"      </one:OE>"+
"    </one:OEChildren>"+
"  </one:Outline>"+
"</one:Page>";
 

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
        }
        [Test]
        public void ProcessSmartTagAndLinkToOneNotePage()
        {
            var smartTag = smartTagAugmenter.GetUnProcessedSmartTags(this.pageContent).First();
            Assert.That(!smartTag.IsProcessed());
            smartTagAugmenter.AddToModel(smartTag,this.pageContent);
            Assert.That(smartTag.IsProcessed());

            smartTagAugmenter.AddLinkToSmartTag(smartTag, this.pageContent, this.defaultSection,"DeadPage"); 
        }
        [Test]
        public void ProcessSmartTagAndLinkToURI()
        {
            var smartTag = smartTagAugmenter.GetUnProcessedSmartTags(this.pageContent).First();
            Assert.That(!smartTag.IsProcessed());
            smartTagAugmenter.AddToModel(smartTag,this.pageContent);
            Assert.That(smartTag.IsProcessed());

            // Set link to URI.
            smartTagAugmenter.AddLinkToSmartTag(smartTag,this.pageContent, new Uri("http://helloworld"));
        }


        [TearDown]
        public void TearDown()
        {
            smartTagNoteBook.Dispose();
            _templateNotebook.Dispose();
        }
    }
}
