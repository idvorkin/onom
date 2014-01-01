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

var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
                     .PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);   
					 
var today = sectionForDailyPages.Page.First(p=>p.name == settings.TodayPageTitle);
ona.OneNoteApplication.NavigateTo(today.ID);