<Query Kind="Statements">
  <Reference Relative="..\..\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll</Reference>
  <Namespace>OneNoteObjectModel</Namespace>
</Query>

var ona = new OneNoteObjectModel.OneNoteApp();
var blankSections = ona.GetNotebooks().Notebook.SelectMany(n=>n.PopulatedSections(ona)).Where(s=>s.IsDefaultUnmodified(ona));
blankSections.ToList().ForEach(s=>ona.OneNoteApplication.DeleteHierarchy(s.ID));

// You can also find and erase blank pages, but that's slow: 
//var blankPages = ona.GetNotebooks().Notebook.SelectMany(n=>n.PopulatedSections(ona).SelectMany(s=>s.Page)).Where(s=>s.IsDefaultUnmodified(ona));
//blankPages.ToList().ForEach(p=>ona.OneNoteApplication.DeleteHierarchy(p.ID));