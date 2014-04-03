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
List<string> OneNoteTextTolist(IEnumerable<Object>  objects)
{
	return objects.SelectMany(o=>
	{
	
		if (o is OEChildren)
		{
			return OneNoteTextTolist((o as OEChildren).Items);
		}
		if (o is OE)
		{
			return OneNoteTextTolist((o as OE).Items);
		}
		if (o is TextRange)
		{
			return 	new List<string>() { (o as TextRange).Value};
		}
		if (o is OneNoteObjectModel.Table)
		{
				// "skipping Table Processing".Dump();
				return new List<string>();
		}
		if (o is OneNoteObjectModel.InkWord)
		{
					return 	new List<string>() { (o as InkWord).recognizedText ?? ""};
		}
		else // (o is OneNoteObjectModel.List)
		{
			o.Dump();
			return new List<string>();
		}
	}
	).ToList();
}
List<string> OneNoteListTolist(IEnumerable<Object>  objects)
{
	return objects.SelectMany(o=>
	{
	
		if (o is OEChildren)
		{
			return OneNoteListTolist((o as OEChildren).Items);
		}
		if (o is OE)
		{
			return OneNoteListTolist((o as OE).Items);
		}
		if (o is TextRange)
		{
			return 	new List<string>() { (o as TextRange).Value};
		}
		else // (o is OneNoteObjectModel.List)
		{
			o.Dump();
			return new List<string>();
		}
	}
	).ToList();
}

Dictionary<string, List<string>> GetTableAfterTitle(string title,IEnumerable<OE> oes)
{
	// TODO test if title isn't there.
	// if (oes.Count
	var table = oes.SkipWhile(i=> {
		var oe = i.Items[0];
		if (!(oe  is TextRange)) return true;
		var text = oe as TextRange;
		return text.Value != title;
	}).Skip(1).First().Items[0] as OneNoteObjectModel.Table;
	return GetTwoColumnTable(table);
}

Dictionary<string, List<string>> GetTwoColumnTable(OneNoteObjectModel.Table table)
{
	
	if (!table.Row.All(r=> r.Cell.Count() == 2))
	{
		throw new ArgumentOutOfRangeException("Table is not 2 dimensional");
	}
	
	// TBD  HANDLE MORE THEN One OEChildren.
	// var c = 
	var keys = table.Row.Select(r=>OneNoteListTolist(r.Cell[0].OEChildren)).Select(k=>k.First());
	var values = table.Row.Select(r=>OneNoteListTolist(r.Cell[1].OEChildren));
	return keys.Zip(values, (k,v) => new {k,v}).ToDictionary(z=>z.k, z=>z.v);
}
IEnumerable<string> GetTableRowContent(OneNoteObjectModel.Page page, string tableTitle, string rowTitle)
{
	var content = ona.GetPageContent(page);
	// ASSUME: Outline is the outer box which contains child items.
	// ASSUME: OEChildren is always a list of OE
	var  children =  content.Items.OfType<Outline>().SelectMany(x=>x.OEChildren);
	var oes = children.SelectMany (x=>x.Items).Cast<OE>();
	return GetTableAfterTitle(tableTitle, oes)[rowTitle];
}
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
	GetTableAfterTitle("SUMMARY", oes).Dump("Summary");
	GetTableAfterTitle("WORK", oes).Dump("Work");
	GetTableAfterTitle("HOME", oes).Dump("Home");
	GetTableAfterTitle("STATS", oes).Dump("Stats");
}

IEnumerable<Tuple<string, string>>  GetTableRowContent (IEnumerable<OneNoteObjectModel.Page> pages, string tableTitle, string rowTitle)
{  
	return pages.SelectMany(  p =>GetTableRowContent(p,tableTitle,rowTitle).Select(lesson => Tuple.Create(p.name, lesson)));
}

IEnumerable<Tuple <string, string>> GetSmartTags(IEnumerable<Object> oes)
{
	var text = OneNoteTextTolist(oes);
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
	
	
	var days = Enumerable.Range(0,8).Select(i=> (DateTime.Now - TimeSpan.FromDays(i)).ToShortDateString());
	var pages = sectionForDailyPages.Page.Where(p=> days.Contains(p.name));
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