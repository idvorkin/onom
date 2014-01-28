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

OneNoteApp ona = new OneNoteApp();
Settings settings = new Settings();
void ViewTemplateNotebook()
{
	Console.WriteLine(settings.DailyPagesNotebook);
}

Button CreateButton(string Content, Action OnClick)
   {
       var button = new Button() { Content = Content };
       button.Click += (o, e) => { OnClick(); };

       button.FontSize = 20;
       return button;
   }

void Main()
{
	var buttons = new []{
		CreateButton("Template Notebook", ViewTemplateNotebook),
		}.ToList();
	buttons.ForEach((b)=>PanelManager.StackWpfElement(b));
}
