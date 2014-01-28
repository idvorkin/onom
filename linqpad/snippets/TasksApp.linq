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
	public string TemplateWeeklyPageTitle = "Weekly";
	public string DailyPagesNotebook = "BlogContentAndResearch";
	public string DailyPagesSection = "Current";
	// TodayPageTitle needs to be a functor as it depends on the day. 
	public Func<String>  TodayPageTitle = () => DateTime.Now.Date.ToShortDateString();
	public Func<String>  ThisWeekPageTitle = () => 	"Week "+ (DateTime.Now.Date - TimeSpan.FromDays((int)DateTime.Now.DayOfWeek-1)).ToShortDateString();
}


public Settings settings = new Settings();
public OneNoteApp ona = new OneNoteObjectModel.OneNoteApp();
	
// Goto page for today, create from template if it doesn't exist yet. 
void GotoTodayPage()
{
	var pageTemplateForDay = ona.GetNotebooks().Notebook.First(n=>n.name == settings.TemplateNotebook )
						.PopulatedSections(ona).First(s=>s.name == settings.TemplateSection)
						.Page.First(p=>p.name == settings.TemplateDailyPageTitle);
	
	var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
	
	if (sectionForDailyPages.Page.Any(p=>p.name == settings.TodayPageTitle()))
	{
		Console.WriteLine("Today's template ({0}) has already been created,going to it",settings.TodayPageTitle);
	}
	else
	{
		var todaysPage = ona.ClonePage(sectionForDailyPages,pageTemplateForDay,settings.TodayPageTitle());
		Console.WriteLine("Created today's template page ({0}).",settings.TodayPageTitle);
		
		// Indent page because it will be folded into a weekly template.
		todaysPage.pageLevel = "2";
		ona.UpdatePage(todaysPage);
		
		// reload section since we modified the tree. 
		sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);    
	}

	var today = sectionForDailyPages.Page.First(p=>p.name == settings.TodayPageTitle());
	ona.OneNoteApplication.NavigateTo(today.ID);	

}

// TODO: Merge with Daily Page
void GotoThisWeekPage()
{
	var pageTemplateForWeek = ona.GetNotebooks().Notebook.First(n=>n.name == settings.TemplateNotebook )
						.PopulatedSections(ona).First(s=>s.name == settings.TemplateSection)
						.Page.First(p=>p.name == settings.TemplateWeeklyPageTitle);
	
	var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
	
	if (sectionForDailyPages.Page.Any(p=>p.name == settings.ThisWeekPageTitle()))
	{
		Console.WriteLine("This Week's template ({0}) has already been created,going to it",settings.TodayPageTitle);
	}
	else
	{
		var thisWeeksPage = ona.ClonePage(sectionForDailyPages,pageTemplateForWeek,settings.ThisWeekPageTitle());
		Console.WriteLine("Creating this week's template page ({0}).",settings.TodayPageTitle);
		ona.UpdatePage(thisWeeksPage);
		
		// reload section since we modified the tree. 
		sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
						.PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);    
	}

	var thisWeek = sectionForDailyPages.Page.First(p=>p.name == settings.ThisWeekPageTitle());
	ona.OneNoteApplication.NavigateTo(thisWeek.ID);	
}

// Remove accidently created empty sections
void DeleteEmptySections()
{
	var ona = new OneNoteObjectModel.OneNoteApp();
	var blankSections = ona.GetNotebooks().Notebook.SelectMany(n=>n.PopulatedSections(ona)).Where(s=>s.IsDefaultUnmodified(ona));
	blankSections.ToList().ForEach(s=>ona.OneNoteApplication.DeleteHierarchy(s.ID));
}

// UX Helpers - These should move to an alternate assembly
Button CreateButton(string Content, Action OnClick)
{
	var button = new Button(){Content = Content};
	button.Click += (o,e) => {OnClick();};
	button.FontSize=20;
	return button;
}

void Main()
{
	var buttons = new []{
		CreateButton("Goto Today", GotoTodayPage),
		CreateButton("Goto This Week", GotoThisWeekPage),
		CreateButton("Erase Empty Section", DeleteEmptySections),
	}.ToList();
	buttons.ForEach((b)=>PanelManager.StackWpfElement(b));
}
