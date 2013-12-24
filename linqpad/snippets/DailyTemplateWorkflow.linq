<Query Kind="Statements">
  <Reference Relative="..\..\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll</Reference>
  <Namespace>OneNoteObjectModel</Namespace>
</Query>

var ona = new OneNoteObjectModel.OneNoteApp();

var pageTemplateForDay = ona.GetNotebooks().Notebook.First(n=>n.name == "Templates")
                      .PopulatedSections(ona).First(s=>s.name == "Default")
                      .Page.First(p=>p.name == "Daily");

var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == "BlogContentAndResearch")
                     .PopulatedSections(ona).First(s=>s.name == "Current");     

var newPageTitle = DateTime.Now.Date.ToShortDateString();
if (sectionForDailyPages.Page.Any(p=>p.name == newPageTitle))
{
    Console.WriteLine("Today's template ({0}) has already been created.",newPageTitle);
}
else
{
    var todaysPage = ona.ClonePage(sectionForDailyPages,pageTemplateForDay,newPageTitle);
    Console.WriteLine("Created today's template page ({0}).",newPageTitle);
	
	// TODO - Reimplement with helper functions
	// todaysPage.pageLevel = 1;
	// ona.OneNoteApplication.UpdatePageContent(todaysPage.ID,
}
