namespace OneNoteObjectModelTests
{
    public class SmartTagTestsPageConent:IPageContentAsText
    {

        private static string first = " <one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">";
        private static string rest = "  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
"  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
"  <one:PageSettings RTL=\"false\" color=\"automatic\">"+
"    <one:PageSize>"+
"      <one:Automatic />"+
"    </one:PageSize>"+
"    <one:RuleLines visible=\"false\" />"+
"  </one:PageSettings>"+
"  <one:Title showTime=\"false\" lang=\"en-US\">"+
"    <one:OE author=\"Igor Dvorkin\" authorInitials=\"ID\" authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-11T18:23:30.000Z\" lastModifiedTime=\"2014-08-11T18:23:30.000Z\" alignment=\"left\" quickStyleIndex=\"0\">"+
"      <one:T><![CDATA[Scratch]]></one:T>"+
"    </one:OE>"+
"  </one:Title>"+
"  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-08-11T18:27:49.000Z\">"+
"    <one:Position x=\"36.0\" y=\"86.4000015258789\" z=\"0\" />"+
"    <one:Size width=\"162.2823486328125\" height=\"40.28314971923828\" />"+
"    <one:OEChildren>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-11T18:23:20.000Z\" lastModifiedTime=\"2014-08-11T18:23:30.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:T><![CDATA[<a"+
" href=\"onenote:SmartTagStorage.one#Model%20981eff41-0428-41f4-8f0c-3ce1ea3546e9&amp;section-id={D6DF7A74-EF20-476B-AC90-8D6D6BBBE4DC}&amp;page-id={EEB32202-CAF4-41B7-AF5A-4BB30DD4327E}&amp;end&amp;extraId={3EF0B7D7-1857-0633-04EE-212D80860CA2}{1}{E19468513483745559900720152138099127754830691}&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">#</a><a"+
" href=\"onenote:Topics.one#processedTag&amp;section-id={25BAB527-01EA-417E-B55F-D52F8E5FB2EC}&amp;page-id={A28B1AD1-4BC0-474A-ABB3-40575F5C6083}&amp;end&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">processedTag</a> is Processed]]></one:T>"+
"      </one:OE>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-11T18:27:49.000Z\" lastModifiedTime=\"2014-08-11T18:27:49.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:T><![CDATA[#unProcessedTag is not ]]></one:T>"+
"      </one:OE>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-11T18:23:30.000Z\" lastModifiedTime=\"2014-08-11T18:23:30.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
"        <one:T><![CDATA[<span"+
"style='text-decoration:line-through'>Text2 is marked through</span>]]></one:T>"+
"      </one:OE>"+
"    </one:OEChildren>"+
"  </one:Outline>"+
"</one:Page>";

        public string firstLine()
        {
            return first;
        }

        public string restOfLines()
        {
            return rest;
        }
    }
}