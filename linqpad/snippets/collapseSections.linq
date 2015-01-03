<Query Kind="Program">
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

public SettingsDailyPages settings = new SettingsDailyPages();
public OneNoteApplication ona = new OneNoteObjectModel.OneNoteApplication();
public ContentProcessor contentProcessor = new ContentProcessor(false);

 /*
void DumpYesterdayPage()
{
	var sectionForDailyPages = DailyPagesSection();
	var today = sectionForDailyPages.Page.First(p=>p.name == (DateTime.Now.Date - TimeSpan.FromDays(1)).ToShortDateString());
	var content = ona.GetPageContent(today);
	// ASSUME: Outline is the outer box which contains child items.
	// ASSUME: OEChildren is always a list of OE
	var  children =  content.Items.OfType<Outline>().SelectMany(x=>x.OEChildren);
	var oes = children.SelectMany (x=>x.Items).Cast<OE>();
	"STATS;SUMMARY".Split(';').ToList()
	.ForEach(key=> contentProcessor.PropertyBagFromTwoColumnTable(contentProcessor.GetTableAfterTitle(key, oes)).Properties.Dump(key));
}
*/
void Main()
{
	var s =  ona.GetNotebook(settings.DailyPagesNotebook).PopulatedSection("Diary 2014");     
	var weekPages = s.Page.Where(p => p.name.StartsWith("Week "));
	GC.KeepAlive(ona);
	GC.SuppressFinalize(ona);
	weekPages.Count().Dump();
	foreach(var p in weekPages)
	{
		p.isCollapsed.Dump();
	//	p.isCollapsed=true;
	//	ona.UpdatePage(p);
	}
}