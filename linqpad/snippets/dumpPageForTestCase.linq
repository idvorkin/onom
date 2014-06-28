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
var page = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook).PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection).Page.First(p=>p.name == "Scratch");
string pageContent;
ona.OneNoteApplication.GetPageContent(page.ID, out pageContent);
var pageContentAsXML = XDocument.Parse(pageContent);
pageContentAsXML.Dump();

// Strip ObjectId's as we're creating new pages.
pageContentAsXML.DescendantNodes() .OfType<XElement>() .ToList() .ForEach(x => x.Attributes().Where(a => a.Name == "objectID").Remove());

var xml = pageContentAsXML.ToString();
Console.Write ("var testPageContent = ");
foreach (var line in xml.Split("\r\n".ToCharArray()).Where(l=>l !=""))
{
	var escapedLine = line.Replace("\"", "\\\"");
	Console.WriteLine("\""+escapedLine+"\"+");
}
Console.WriteLine(";");