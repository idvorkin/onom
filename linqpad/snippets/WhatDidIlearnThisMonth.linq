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

IEnumerable<Tuple<string, string>>  GetTableRowContent (IEnumerable<OneNoteObjectModel.Page> pages, string tableTitle, string rowTitle)
{  	
	return pages.SelectMany(  p =>contentProcessor.GetTableRowContent(p,tableTitle,rowTitle).Select(lesson => Tuple.Create(p.name, lesson)));
}

void WhatDidILearnLastMonth()
{
	var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
	
	
	var pages = sectionForDailyPages.Page.Where(p=> p.name.StartsWith("Week ")).ToList();
	// GetTableRowContent(pages,"SUMMARY","What did I learn").Dump("What did I learn last week");
	var propBags = pages.Select ( week =>
	{
		var oes = ona.GetPageContent(week).Items.OfType<Outline>().SelectMany(x => x.OEChildren).SelectMany(x => x.Items).Cast<OE>();
		var summaryTable = contentProcessor.GetTableAfterTitle("SUMMARY",oes);
		var whatDidILearnTable = summaryTable.Row.Select(r=>r.Cell[1].OEChildren.First().Items.First()).Cast<OE>().Select(oe=>oe.Items[0]).OfType<Table>().First();
		return contentProcessor.PropertyBagFromOneColumnPropertyTable(whatDidILearnTable);
	});
	new PropertyBag().Merge(propBags).Dump();
}

void Main()
{
	WhatDidILearnLastMonth();
}