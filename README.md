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

Erase Blank Sections and Pages
---------------------
It's easy to accidently create new pages and sections, this script finds and erases them: 

```csharp
// Find and erase blank pages
var ona = new OneNoteObjectModel.OneNoteApp();
var blankSections = ona.GetNotebooks().Notebook.SelectMany(n=>n.PopulatedSections(ona)).Where(s=>s.IsDefaultUnmodified(ona));
blankSections.ToList().ForEach(s=>ona.OneNoteApplication.DeleteHierarchy(s.ID));

// You can also find and erase blank pages, but that's slow: 
var blankPages = ona.GetNotebooks().Notebook.SelectMany(n=>n.PopulatedSections(ona).SelectMany(s=>s.Page)).Where(s=>s.IsDefaultUnmodified(ona));
blankPages.ToList().ForEach(p=>ona.OneNoteApplication.DeleteHierarchy(p.ID));
```

Erase Blank Pages
---------------------
It's easy to accidently create a new section, this script finds and erases them: 

```csharp
var ona = new OneNoteObjectModel.OneNoteApp();
var blankSections = ona.GetNotebooks().Notebook.SelectMany(n=>n.PopulatedSections(ona)).Where(s=>s.IsDefaultUnmodified(ona));
blankSections.ToList().ForEach(s=>ona.OneNoteApplication.DeleteHierarchy(s.ID));
```


Erase Blank Pages
-------------------


Create a new page for every day from a template:
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

