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

public SettingsDailyPages settings = new SettingsDailyPages();
public OneNoteApp ona = new OneNoteObjectModel.OneNoteApp();
public ContentProcessor contentProcessor = new ContentProcessor(new OneNoteObjectModel.OneNoteApp(),false);

IEnumerable<Tuple <string, string>> GetSmartTags(IEnumerable<Object> oes)
{
	var text = contentProcessor.OneNoteContentToList(oes); //5.Where(t=>t != null);
	var smartTagMatcher = new Regex("(#[a-zA-Z]+)");
	return text.Select( t => new {t, match=smartTagMatcher.Match(t)}).Where(x => x.match.Success).Select(x=>
	{
		var tag = x.match.Captures[0].Value;
		return Tuple.Create(tag, smartTagMatcher.Replace(x.t,""));
	});
}

void Main()
{
	// 1) Find Current Page
	var pageId = ona.OneNoteApplication.Windows.CurrentWindow.CurrentPageId;
	string pageDataXML;
	ona.OneNoteApplication.GetPageContent(pageId,out pageDataXML);
	var pageData = OneNoteApp.XMLDeserialize<OneNoteObjectModel.Page>(pageDataXML);
	
	var  oes =  pageData.Items.OfType<Outline>().SelectMany(x=>x.OEChildren).SelectMany(x=>x.Items).Cast<OE>();
	GetSmartTags(oes).Dump("Smart Tags for Page:"+pageData.name);
}