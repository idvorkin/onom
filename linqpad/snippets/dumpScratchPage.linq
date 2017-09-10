<Query Kind="Statements">
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OnenoteCapabilities\bin\Debug\OnenoteCapabilities.dll">C:\gits\onom\OnenoteCapabilities\bin\Debug\OnenoteCapabilities.dll</Reference>
  <Reference Relative="..\..\OnenoteCapabilities\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OnenoteCapabilities\bin\Debug\OneNoteObjectModel.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\PresentationCore.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\PresentationFramework.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\System.Windows.Presentation.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Xaml.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\WindowsBase.dll</Reference>
  <Namespace>OnenoteCapabilities</Namespace>
  <Namespace>OneNoteObjectModel</Namespace>
  <Namespace>System.Windows</Namespace>
  <Namespace>System.Windows.Controls</Namespace>
</Query>

SettingsDailyPages settings = new SettingsDailyPages();
var ona = new OneNoteApplication();

var ContentProcessor = new ContentProcessor();


var page = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook).PopulatedSections().First(s=>s.name == settings.DailyPagesSection).Page.First(p=>p.name == "Scratch");
var rowTemplate = "<one:Row lastModifiedTime=\"2014-06-28T06:11:19.000Z\" xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"> " +
  "<one:Cell lastModifiedTime=\"2014-06-28T06:11:19.000Z\"  lastModifiedByInitials=\"ID\"> " +
    "<one:OEChildren> " +
      "<one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:11:19.000Z\" lastModifiedTime=\"2014-06-28T06:11:19.000Z\" alignment=\"left\" quickStyleIndex=\"1\"> " +
        "<one:Tag index=\"0\" completed=\"{2}\" disabled=\"false\" creationDate=\"2014-06-28T06:11:25.000Z\" completionDate=\"2014-06-28T06:27:01.000Z\" />"+
		"<one:T><![CDATA[{0}]]></one:T> " +
      "</one:OE> " +
    "</one:OEChildren> " +
  "</one:Cell> " +
  "<one:Cell lastModifiedTime=\"2014-06-28T06:11:13.000Z\" lastModifiedByInitials=\"ID\"> " +
    "<one:OEChildren> " +
      "<one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:11:13.000Z\" lastModifiedTime=\"2014-06-28T06:11:13.000Z\" alignment=\"left\" quickStyleIndex=\"1\"> " +
      "<one:T><![CDATA[{1}]]></one:T> " +
      "</one:OE> " +
    "</one:OEChildren> " +
  "</one:Cell> " +
"</one:Row>";

bool completed=false;
var row = string.Format(rowTemplate,"taskName",DateTime.Now,completed.ToString().ToLower());

var rowAsXML = XDocument.Parse(row);
rowAsXML.Dump();

string pageContent;
var c = ona.GetPageContentAsXDocument(page);
c.Dump("Page Content");


var content = ona.GetPageContent(page);
content.Dump();

//pageContentAsXML.Dump();
//pageContentAsXML.ToString().Dump();
//pageContentAsXML.DescendantNodes().OfType<XElement>().Where(e=>e.Name.LocalName=="Row").Last().Dump();
//pageContentAsXML.DescendantNodes().OfType<XElement>().Where(e=>e.Name.LocalName=="Row").First().AddAfterSelf(rowAsXML.Root);
//pageContentAsXML.Dump();
//ona.OneNoteApplication.UpdatePageContent(pageContentAsXML.ToString());