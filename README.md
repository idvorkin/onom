Onenote Object Model (onom)
====

The OneNote Object Model exposed via C#. This project contains the following sub directories:

* OneNoteObjectModel - An assembly exposing the ObjectModel
* ListNotebooks - An example showing listing the notebooks. 

Features:
==========

Typed Helpers:
----------------
Many OneNote APIs, like GetHierarchy, are not conveniently typed, so ONOM implements typed helpers, like GetNotebooks, which are exposed on the OneNoteApp class. 

```csharp
// Access your notebooks and sections using Linq

var ona = new OneNoteObjectModel.OneNoteApp();
var notebook = ona.GetNotebooks().Notebook.First(n=>n.name == "BlogContentAndResearch");
var section = ona.GetSections(notebook).First(s=>s.name == "Current");
```

Clone Pages:
----------------
Updating raw onenote page XML is miserable.  Thus, there is an API's for [cloning pages](http://share.linqpad.net/ekfuve.linq):

```csharp

// Clone the daily template into a page named with today's date.
var ona = new OneNoteObjectModel.OneNoteApp();

var templateForDay = ona.GetNotebooks().Notebook.First(n=>n.name == "Templates")
					  .PopulatedSections(ona).First(s=>s.name == "Default")
					  .Page.First(p=>p.name == "Daily");
					  
var sectionForDayPage = ona.GetNotebooks().Notebook.First(n=>n.name == "BlogContentAndResearch")
					 .PopulatedSections(ona).First(s=>s.name == "Current");		

var newPageTitle = DateTime.Now.Date.ToShortDateString();
if (sectionForDayPage.Page.Any(p=>p.name == newPageTitle))
{
	Console.WriteLine("Today's template ({0}) has already been created.",newPageTitle);
}
else
{
	ona.ClonePage(sectionForDayPage,templateForDay,newPageTitle);
	Console.WriteLine("Created today's template page ({0}).",newPageTitle);
}

```


TODO List:

* Create a NuGet project for easy access from other projects and LinqPad

