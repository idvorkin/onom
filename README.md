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

var onom = new OneNoteObjectModel.OneNoteApp();
var notebook = onom.GetNotebooks().Notebook.First(n=>n.name == "BlogContentAndResearch");
var section = onom.GetSections(notebook).First(s=>s.name == "Current");
```

Clone Pages:
----------------
Updating raw onenote page XML is miserable.  Thus, there is an API's for cloning pages:

```csharp
// Clone the daily template into a page named with today's date.

var onom = new OneNoteObjectModel.OneNoteApp();
var notebook = onom.GetNotebooks().Notebook.First(n=>n.name == "BlogContentAndResearch");
var section = onom.GetSections(notebook).First(s=>s.name == "Current");
var template = section.Page.First(p=>p.name == "Daily Template");
var newPageTitle = DateTime.Now.Date.ToShortDateString();
if (section.Page.Any(p=>p.name == newPageTitle))
{
    Console.WriteLine("Today's template has already been created");
}
else
{
    onom.ClonePage(section,template,newPageTitle);
}
```


TODO List:

* Create a NuGet project for easy access from other projects and LinqPad

