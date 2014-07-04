using System;
using System.Linq;
using System.Xml.Linq;
using NUnit.Framework;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    public class DumbTodoTests
    {
        private OneNoteApp ona;

        [SetUp]
        public void Setup()
        {
            this.ona = new OneNoteApp();
            this._scratchNotebook = new TemporaryNoteBookHelper(ona, "dumbTodoScratch");

            // create template structure.
            this.tempSection  = ona.CreateSection(_scratchNotebook.Get(), "dumbTodoTestSection");
            var page = ona.CreatePage(tempSection, "dumpTodoTestPage");

            // the copied page has a PageID, replace it with PageID from the newly created page.
            var filledInPage = string.Format(pageWithTwoDumbTodoTables, page.ID, page.name);
            this.pageContent = XDocument.Parse(filledInPage);
            ona.OneNoteApplication.UpdatePageContent(pageContent.ToString());
        }

        // TODO: Come up with a better template for such tests.
        readonly string pageWithTwoDumbTodoTables = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">"+
                    "  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
                    "  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
                    "  <one:PageSettings RTL=\"false\" color=\"automatic\">"+
                    "    <one:PageSize>"+
                    "      <one:Automatic />"+
                    "    </one:PageSize>"+
                    "    <one:RuleLines visible=\"false\" />"+
                    "  </one:PageSettings>"+
                    "  <one:Title lang=\"en-US\">"+
                    "    <one:OE author=\"Igor Dvorkin\" authorInitials=\"ID\" authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T22:20:09.000Z\" lastModifiedTime=\"2014-06-28T22:20:09.000Z\" alignment=\"left\" quickStyleIndex=\"0\">"+
                    "      <one:T><![CDATA[Scratch]]></one:T>"+
                    "    </one:OE>"+
                    "  </one:Title>"+
                    "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-07-02T13:10:11.000Z\">"+
                    "    <one:Position x=\"54.0\" y=\"122.400001525879\" z=\"1\" />"+
                    "    <one:Size width=\"72.21826934814453\" height=\"13.42771339416504\" />"+
                    "    <one:OEChildren>"+
                    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-02T13:10:11.000Z\" lastModifiedTime=\"2014-07-02T13:10:11.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "        <one:T><![CDATA[Table 1]]></one:T>"+
                    "      </one:OE>"+
                    "    </one:OEChildren>"+
                    "  </one:Outline>"+
                    "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-07-02T13:10:09.000Z\">"+
                    "    <one:Position x=\"36.0\" y=\"176.4000091552734\" z=\"0\" />"+
                    "    <one:Size width=\"90.53804016113281\" height=\"40.35544204711914\" />"+
                    "    <one:OEChildren>"+
                    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:32:37.000Z\" lastModifiedTime=\"2014-07-02T13:10:07.000Z\" alignment=\"left\">"+
                    "        <one:Table bordersVisible=\"true\" hasHeaderRow=\"true\" lastModifiedTime=\"2014-07-02T13:10:07.000Z\">"+
                    "          <one:Columns>"+
                    "            <one:Column index=\"0\" width=\"37.11000061035156\" />"+
                    "            <one:Column index=\"1\" width=\"45.29803085327148\" />"+
                    "          </one:Columns>"+
                    "          <one:Row lastModifiedTime=\"2014-07-02T13:10:07.000Z\">"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:07.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:32:36.000Z\" lastModifiedTime=\"2014-07-02T13:10:07.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
                    "                  <one:T><![CDATA[<span"+
                    "style='font-weight:bold'>Task</span>]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:07.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:33:45.000Z\" lastModifiedTime=\"2014-07-02T13:10:07.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
                    "                  <one:T><![CDATA[<span"+
                    "style='font-weight:bold'>Due Date</span>]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "          </one:Row>"+
                    "          <one:Row lastModifiedTime=\"2014-07-02T13:10:07.000Z\">"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:07.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:33:24.000Z\" lastModifiedTime=\"2014-07-02T13:10:07.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "                  <one:T><![CDATA[]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:07.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-01T22:20:49.000Z\" lastModifiedTime=\"2014-07-02T13:10:07.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "                  <one:T><![CDATA[]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "          </one:Row>"+
                    "        </one:Table>"+
                    "      </one:OE>"+
                    "    </one:OEChildren>"+
                    "  </one:Outline>"+
                    "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-07-02T13:39:28.000Z\">"+
                    "    <one:Position x=\"342.0\" y=\"212.4000091552734\" z=\"3\" />"+
                    "    <one:Size width=\"72.0\" height=\"13.42771339416504\" />"+
                    "    <one:OEChildren indent=\"2\">"+
                    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-02T13:34:03.000Z\" lastModifiedTime=\"2014-07-02T13:34:03.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "        <one:T><![CDATA[]]></one:T>"+
                    "      </one:OE>"+
                    "    </one:OEChildren>"+
                    "  </one:Outline>"+
                    "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\">"+
                    "    <one:Position x=\"270.0\" y=\"284.3999938964844\" z=\"2\" />"+
                    "    <one:Size width=\"90.53805541992187\" height=\"67.21086883544922\" />"+
                    "    <one:OEChildren>"+
                    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-02T13:10:14.000Z\" lastModifiedTime=\"2014-07-02T13:10:14.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "        <one:T><![CDATA[Table 2]]></one:T>"+
                    "      </one:OE>"+
                    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:32:37.000Z\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\" alignment=\"left\">"+
                    "        <one:Table bordersVisible=\"true\" hasHeaderRow=\"true\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\">"+
                    "          <one:Columns>"+
                    "            <one:Column index=\"0\" width=\"37.11000061035156\" />"+
                    "            <one:Column index=\"1\" width=\"45.29803085327148\" />"+
                    "          </one:Columns>"+
                    "          <one:Row lastModifiedTime=\"2014-07-02T13:10:24.000Z\">"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:24.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:32:36.000Z\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
                    "                  <one:T><![CDATA[<span"+
                    "style='font-weight:bold'>Task</span>]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:24.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:33:45.000Z\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
                    "                  <one:T><![CDATA[<span"+
                    "style='font-weight:bold'>Due Date</span>]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "          </one:Row>"+
                    "          <one:Row lastModifiedTime=\"2014-07-02T13:10:24.000Z\">"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:24.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:33:24.000Z\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "                  <one:T><![CDATA[]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "            <one:Cell lastModifiedTime=\"2014-07-02T13:10:24.000Z\" lastModifiedByInitials=\"ID\">"+
                    "              <one:OEChildren>"+
                    "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-01T22:20:49.000Z\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "                  <one:T><![CDATA[]]></one:T>"+
                    "                </one:OE>"+
                    "              </one:OEChildren>"+
                    "            </one:Cell>"+
                    "          </one:Row>"+
                    "        </one:Table>"+
                    "      </one:OE>"+
                    "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-02T13:10:24.000Z\" lastModifiedTime=\"2014-07-02T13:10:24.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                    "        <one:T><![CDATA[]]></one:T>"+
                    "      </one:OE>"+
                    "    </one:OEChildren>"+
                    "  </one:Outline>"+
                    "</one:Page>";

        private XDocument pageContent;
        private TemporaryNoteBookHelper _scratchNotebook;
        private Section tempSection;

        [Test]
        public void TestPageTemplateIsValid()
        {
            // If this test passes our template is correctly making pages.
        }

        [Test]
        public void AddDumbTodo()
        {

            // Write to table 0. 
            DumbTodo.AddToPage(ona, pageContent, "this is a todo", null, 0);
            DumbTodo.AddToPage(ona, pageContent, "this is a todo", DateTime.Today, 0);



            // write to table 1
            DumbTodo.AddToPage(ona, pageContent, "this is a todo", null, 1);
            DumbTodo.AddToPage(ona, pageContent, "this is a todo", DateTime.Today, 1);


            // TODAY: Assert not crashing 
            // future test todos actually get written.
        }

        [Test]
        public void AddDumbTodoFromSmartTags()
        {
            var smartTag = new SmartTag()
            {
                FullText = "#hello this is a todo"
            };

            var smartTagWithDate = new SmartTag()
            {
                FullText = "#hello this is a todo next week"
            };



            // Write smartTag.
            DumbTodo.AddToPageFromDateEnableSmartTag(ona, pageContent, smartTag, 0);
            DumbTodo.AddToPageFromDateEnableSmartTag(ona, pageContent, smartTag, 1);


            // write smartTag with date.
            DumbTodo.AddToPageFromDateEnableSmartTag(ona, pageContent, smartTag, 0);
            DumbTodo.AddToPageFromDateEnableSmartTag(ona, pageContent, smartTag, 1);


            // TODAY: Assert not crashing 
            // future test todos actually get written.
        }




        [TearDown]
        public void TearDown()
        {
            _scratchNotebook.Dispose();
        }
    }
}