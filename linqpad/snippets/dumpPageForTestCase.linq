<Query Kind="Statements">
  <Reference Relative="..\..\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OnenoteCapabilities\bin\Debug\OnenoteCapabilities.dll">C:\gits\onom\OnenoteCapabilities\bin\Debug\OnenoteCapabilities.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll</Reference>
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
OneNoteApp ona = new OneNoteObjectModel.OneNoteApp();
ContentProcessor contentProcessor = new ContentProcessor(new OneNoteObjectModel.OneNoteApp(),false);
var page = ona.GetNotebook(settings.DailyPagesNotebook).PopulatedSection(ona,settings.DailyPagesSection)
.GetPage(ona,"Scratch");
string pageContent;
ona.OneNoteApplication.GetPageContent(page.ID, out pageContent);
var pageContentAsXML = XDocument.Parse(pageContent);
pageContentAsXML.Dump();

// Strip ObjectId's as we're creating new pages.
pageContentAsXML.DescendantNodes() .OfType<XElement>() .ToList() .ForEach(x => x.Attributes().Where(a => a.Name == "objectID").Remove());

var xml = pageContentAsXML.ToString();
Console.Write ("var testPageContent = ");

// First line needs template replacment for ID={0} and PageNamge={1}
var firstLine = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">";

var lines = new List<string> () {firstLine};
lines.AddRange(xml.Split("\r\n".ToCharArray()).Skip(1).Where(l=>l !=""));

foreach (var line in lines)
{
	var escapedLine = line.Replace("\"", "\\\"");
	Console.WriteLine("\""+escapedLine+"\"+");
}
Console.WriteLine(";");