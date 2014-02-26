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
List<string> OneNoteTextTolist(IEnumerable<Object>  objects)
{
	return objects.SelectMany(o=>
	{
	
		if (o is OEChildren)
		{
			return OneNoteTextTolist((o as OEChildren).Items);
		}
		if (o is OE)
		{
			return OneNoteTextTolist((o as OE).Items);
		}
		if (o is TextRange)
		{
			return 	new List<string>() { (o as TextRange).Value};
		}
		if (o is OneNoteObjectModel.Table)
		{
				// "skipping Table Processing".Dump();
				return new List<string>();
		}
		if (o is OneNoteObjectModel.InkWord)
		{
			var ink = o as InkWord;
			return 	new List<string>() { (o as InkWord).recognizedText ?? ""};
		}
		else // (o is OneNoteObjectModel.List)
		{
			o.Dump();
			return new List<string>();
		}
	}
	).ToList();
}

IEnumerable<Tuple <string, string>> GetSmartTags(IEnumerable<Object> oes)
{
	var text = OneNoteTextTolist(oes); //5.Where(t=>t != null);
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