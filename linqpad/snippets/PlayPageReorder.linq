<Query Kind="Statements">
  <Reference Relative="..\..\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll">C:\gits\onom\InterOpAssembly\Microsoft.Office.Interop.OneNote.dll</Reference>
  <Reference Relative="..\..\OnenoteCapabilities\bin\Debug\OnenoteCapabilities.dll">C:\gits\onom\OnenoteCapabilities\bin\Debug\OnenoteCapabilities.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll">C:\gits\onom\OneNoteObjectModel\bin\Debug\OneNoteObjectModel.dll</Reference>
  <Reference Relative="..\..\OneNoteObjectModelTests\bin\Debug\OneNoteObjectModelTests.dll">C:\gits\onom\OneNoteObjectModelTests\bin\Debug\OneNoteObjectModelTests.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\PresentationCore.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\PresentationFramework.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\System.Windows.Presentation.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Xaml.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\WPF\WindowsBase.dll</Reference>
  <Namespace>OnenoteCapabilities</Namespace>
  <Namespace>OneNoteObjectModel</Namespace>
  <Namespace>OneNoteObjectModelTests</Namespace>
  <Namespace>System.Windows</Namespace>
  <Namespace>System.Windows.Controls</Namespace>
</Query>

/*
Expose useful onenote capability in a GUI application hosted in linqpad. 

To Use: 
   Update settings with your notebook names
   Run the script, and click the buttons under the custom tab
   
To Add new functionality:
   Add a function and then call it from a button.

As functionality matures
	Move to Onenote Capabilities
*/



var ona = new OneNoteApp();
using (var nbh = new TemporaryNoteBookHelper(ona))
{
	var notebook =  nbh.Get();
    var tempSection  = ona.CreateSection(nbh.Get(), "random");
    ona.CreatePage(tempSection,"A");
    ona.CreatePage(tempSection, "B");
    ona.CreatePage(tempSection, "C");
	ona.CreatePage(tempSection, "D");


	var section = ona.GetSections(notebook,true).First();
	section.Page.Select(p=>p.name).Dump();

	var pages = section.Page.ToList();
	var lastElement = pages.Last();
	pages.Insert(pages.IndexOf(pages.First(p=>p.name=="A"))+1,lastElement);
	pages.Dump();
	section.Page = pages.Take(pages.Count - 1).ToArray();

	ona.OneNoteApplication.UpdateHierarchy(OneNoteApp.XMLSerialize<Section>(section));
	
	section = ona.GetSections(notebook,true).First();
	section.Page.Select(p=>p.name).Dump();
}


