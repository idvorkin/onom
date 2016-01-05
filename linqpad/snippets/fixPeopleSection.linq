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


// After a while people section gets messed up
// This will reorder people section alphabetically, and by date, and collapse weeks.
// It also pulls to top out of order items.


void Main()
{
	var s = ona.GetNotebook(settings.DailyPagesNotebook).PopulatedSection("People");
	s.Page.Count().Dump("Total Pages in People Section ");


	var peoplePages = s.Page.Where(p => p.name.Split(':').Count() == 2);
	var nonPeoplePages = s.Page.Where(p => !peoplePages.Contains(p));

	nonPeoplePages.ToList().Select(p => $"{p.name}").Dump("Non People Pages");

	// Group By Person Name
	var groupedPages = peoplePages.GroupBy(p => p.name.Split(':')[0]);
	
	var newPageOrder = new List<OneNoteObjectModel.Page>();
	foreach (var p in nonPeoplePages)
	{
		p.pageLevel=1.ToString();
		newPageOrder.Add(p);	
	}

	//foreach (var p in peoplePages) newPageOrder.Add(p);
	foreach (var gp in groupedPages.OrderBy(gp=>gp.Key))
	{
		var parentPage = gp.FirstOrDefault(p => p.name.Contains(":Next"));
		if (parentPage == null)
		{
			gp.Dump("Broken Section");
			continue;
		}
		parentPage.isCollapsed=true;
		parentPage.pageLevel="1".ToString();
		var childPages = gp.Where(p => p.name != parentPage.name);
		newPageOrder.Add(parentPage);
		foreach (var cp in childPages)
		{
			var date = cp.name.Split(':')[1];
			cp.pageLevel = "2".ToString();
			DateTime dt;
			var parsedDate = DateTime.TryParse(date, out dt);
			if (!parsedDate) gp.Dump("Broken Section");	
		}
		newPageOrder.AddRange(childPages.OrderByDescending(cp=>DateTime.Parse(cp.name.Split(':')[1])));
	}
	
	var isOriginalAndReorderedPageCountMatch = s.Page.Count() == newPageOrder.Count;
	isOriginalAndReorderedPageCountMatch.Dump("Does page order match" );
	newPageOrder.ToList().Select(p => $"{p.name}, {p.pageLevel}").Dump("NewPageOrder");
	"Uncomment next lines to write new page order".Dump();
	// if (!isOriginalAndReorderedPageCountMatch) throw new Exception("Safety");
	//s.Page = newPageOrder.ToArray();
	// OneNoteApplication.Instance.InteropApplication.UpdateHierarchy(OneNoteApplication.XMLSerialize(s));
}