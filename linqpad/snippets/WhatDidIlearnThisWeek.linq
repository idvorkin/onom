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
public OneNoteApp ona = new OneNoteObjectModel.OneNoteApp();
public ContentProcessor contentProcessor = new ContentProcessor(new OneNoteObjectModel.OneNoteApp(),false);
void DumpYesterdayPage()
{
 	var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
	
	var today = sectionForDailyPages.Page.First(p=>p.name == (DateTime.Now.Date - TimeSpan.FromDays(1)).ToShortDateString());
	var content = ona.GetPageContent(today);
	// ASSUME: Outline is the outer box which contains child items.
	// ASSUME: OEChildren is always a list of OE
	var  children =  content.Items.OfType<Outline>().SelectMany(x=>x.OEChildren);
	var oes = children.SelectMany (x=>x.Items).Cast<OE>();
	"STATS;SUMMARY".Split(';').ToList()
	.ForEach(key=> contentProcessor.PropertyBagFromTwoColumnTable(contentProcessor.GetTableAfterTitle(key, oes)).Properties.Dump(key));
}

IEnumerable<Tuple<string, string>>  GetTableRowContent (IEnumerable<OneNoteObjectModel.Page> pages, string tableTitle, string rowTitle)
{  	
	return pages.SelectMany(  p =>contentProcessor.GetTableRowContent(p,tableTitle,rowTitle).Select(lesson => Tuple.Create(p.name, lesson)));
}

IEnumerable<Tuple <string, string>> GetSmartTags(IEnumerable<Object> oes)
{
	var text = contentProcessor.OneNoteContentToList(oes);
	var smartTagMatcher = new Regex("(#[a-zA-Z]+)");
	return text.Select( t => new {t, match=smartTagMatcher.Match(t)}).Where(x => x.match.Success).Select(x=>
	{
		var tag = x.match.Captures[0].Value;
		return Tuple.Create(tag, smartTagMatcher.Replace(x.t,""));
	});
}
void WhatDidILearnLastWeek()
{
	var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
	
	
	var days = Enumerable.Range(0,7).Select(i=> (DateTime.Now - TimeSpan.FromDays(i)).ToShortDateString());
	var pages = sectionForDailyPages.Page.Where(p=> days.Contains(p.name)).ToList();
	GetTableRowContent(pages,"SUMMARY","What did I learn").Dump("What did I learn last week");
	pages.ToList().Select(p=>ona.GetPageContent(p)).Select( p => 
	{
		var  oes =  p.Items.OfType<Outline>().SelectMany(x=>x.OEChildren).SelectMany(x=>x.Items).OfType<OE>();
		return new {p.name, smartTags = GetSmartTags(oes) };
	}
	).Dump("Smart Tags");
}
void Main()
{
	WhatDidILearnLastWeek();
}