namespace OneNoteObjectModelTests
{
    public class SmartTagTestsPageConent:IPageContentAsText
    {

        private static string first = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">";

        private static string rest = "  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />" +
            "  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />" +
            "  <one:PageSettings RTL=\"false\" color=\"automatic\">" +
            "    <one:PageSize>" +
            "      <one:Automatic />" +
            "    </one:PageSize>" +
            "    <one:RuleLines visible=\"false\" />" +
            "  </one:PageSettings>" +
            "  <one:Title lang=\"en-US\">" +
            "    <one:OE author=\"Igor Dvorkin\" authorInitials=\"ID\" authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-25T06:36:27.000Z\" lastModifiedTime=\"2014-07-25T06:36:27.000Z\" alignment=\"left\" quickStyleIndex=\"0\">" +
            "      <one:T><![CDATA[Scratch]]></one:T>" +
            "    </one:OE>" +
            "  </one:Title>" +
            "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-08-03T16:35:33.000Z\">" +
            "    <one:Position x=\"54.0\" y=\"122.400001525879\" z=\"0\" />" +
            "    <one:Size width=\"163.4356689453125\" height=\"40.28313827514648\" />" +
            "    <one:OEChildren>" +
            "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;AD&quot; hash=&quot;h6AIXFvXPSXcdyef4IPVVA==&quot;&gt;&lt;localId sid=&quot;S-1-5-21-2127521184-1604012920-1887927527-1154866&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-03T16:34:45.000Z\" lastModifiedTime=\"2014-08-03T16:35:33.000Z\" alignment=\"left\" quickStyleIndex=\"1\">" +
            "        <one:T><![CDATA[<a" +
            " href=\"onenote:#Model: 037fef74-1ea2-47ae-81e5-2266016aaabb&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch/SmartTagStorage.one&amp;extraId={B686A30C-EB43-04F3-1B67-57E8C30A14ED}{1}{E1955612343907131023571976342452626004540251}\">#</a><a" +
            " href=\"onenote:#ProcessedTag&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch/Topics.one\">ProcessedTag</a> is Processed]]></one:T>" +
            "      </one:OE>" +
            "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;AD&quot; hash=&quot;h6AIXFvXPSXcdyef4IPVVA==&quot;&gt;&lt;localId sid=&quot;S-1-5-21-2127521184-1604012920-1887927527-1154866&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-03T16:35:07.000Z\" lastModifiedTime=\"2014-08-03T16:35:33.000Z\" alignment=\"left\" quickStyleIndex=\"1\">" +
            "        <one:T><![CDATA[#UnProcessedTag is not ]]></one:T>" +
            "      </one:OE>" +
            "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;AD&quot; hash=&quot;h6AIXFvXPSXcdyef4IPVVA==&quot;&gt;&lt;localId sid=&quot;S-1-5-21-2127521184-1604012920-1887927527-1154866&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-03T16:34:55.000Z\" lastModifiedTime=\"2014-08-03T16:35:33.000Z\" alignment=\"left\" quickStyleIndex=\"1\">" +
            "        <one:T><![CDATA[Text2 is marked through]]></one:T>" +
            "      </one:OE>" +
            "    </one:OEChildren>" +
            "  </one:Outline>" +
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