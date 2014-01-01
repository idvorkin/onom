<Query Kind="Statements">
  <Reference Relative="..\..\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll</Reference>
  <Namespace>OneNoteObjectModel</Namespace>
</Query>

var settings = new {
	TemplateNotebook  = "Templates",
	TemplateSection = "Default",
	TemplateDailyPageTitle = "Daily",
	DailyPagesNotebook = "BlogContentAndResearch",
	DailyPagesSection = "Current",
	TodayPageTitle = DateTime.Now.Date.ToShortDateString() 
};


var ona = new OneNoteObjectModel.OneNoteApp();

var pageTemplateForDay = ona.GetNotebooks().Notebook.First(n=>n.name == settings.TemplateNotebook )
                      .PopulatedSections(ona).First(s=>s.name == settings.TemplateSection)
                      .Page.First(p=>p.name == settings.TemplateDailyPageTitle);

var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
                     .PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     

if (sectionForDailyPages.Page.Any(p=>p.name == settings.TodayPageTitle))
{
    Console.WriteLine("Today's template ({0}) has already been created.",settings.TodayPageTitle);
}
else
{
    var todaysPage = ona.ClonePage(sectionForDailyPages,pageTemplateForDay,settings.TodayPageTitle);
    Console.WriteLine("Created today's template page ({0}).",settings.TodayPageTitle);
	
	// Indent page because it will be folded into a weekly template.
	 todaysPage.pageLevel = "2";
	 ona.UpdatePage(todaysPage);

}