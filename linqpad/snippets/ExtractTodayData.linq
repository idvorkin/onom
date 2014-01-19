<Query Kind="Program">
  <Reference Relative="..\..\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\PresentationCore.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\PresentationFramework.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\System.Windows.Presentation.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Xaml.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\WindowsBase.dll</Reference>
  <Namespace>OneNoteObjectModel</Namespace>
  <Namespace>System.Windows</Namespace>
  <Namespace>System.Windows.Controls</Namespace>
</Query>

/*
Expose useful onenote capability in a GUI application hosted in linqpad. 

To Use: 
   Update settings with your notebook names
   Run the script, and click the buttons under the custom tab
   
To Addd new functionality:
   Add a function and then call it from a button.
*/
public class Settings {
	public string TemplateNotebook  = "Templates";
	public string TemplateSection = "Default";
	public string TemplateDailyPageTitle = "Daily";
	public string DailyPagesNotebook = "BlogContentAndResearch";
	public string DailyPagesSection = "Current";
	public string TodayPageTitle = DateTime.Now.Date.ToShortDateString();
};

public Settings settings = new Settings();
public OneNoteApp ona = new OneNoteObjectModel.OneNoteApp();
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
void GotoTodayPage()
{
	var pageTemplateForDay = ona.GetNotebooks().Notebook.First(n=>n.name == settings.TemplateNotebook )
						.PopulatedSections(ona).First(s=>s.name == settings.TemplateSection)
						.Page.First(p=>p.name == settings.TemplateDailyPageTitle);
	
	var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
	
	var today = sectionForDailyPages.Page.First(p=>p.name == "1/17/2014");
	var content = ona.GetPageContent(today);
	
	// Outline is the outer box which contains child items.
	var  children =  content.Items.OfType<Outline>().SelectMany(x=>x.OEChildren);
	GetTableAfterTitle("SUMMARY",children.SelectMany (x=>x.Items).Cast<OE>()).Dump("Summary");
	GetTableAfterTitle("WORK",children.SelectMany (x=>x.Items).Cast<OE>()).Dump("Work");
	GetTableAfterTitle("HOME",children.SelectMany (x=>x.Items).Cast<OE>()).Dump("Home");
	GetTableAfterTitle("STATS",children.SelectMany (x=>x.Items).Cast<OE>()).Dump("Stats");
}
void Main()
{
	GotoTodayPage();
}