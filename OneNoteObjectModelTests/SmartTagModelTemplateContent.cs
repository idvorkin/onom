using System.Runtime.Serialization.Formatters;

namespace OneNoteObjectModelTests
{
    public class SmartTagModelTemplateContent:IPageContentAsText
    {
        internal readonly string first = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">";
        internal readonly string rest = "  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
                                                                              "  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
                                                                              "  <one:PageSettings RTL=\"false\" color=\"automatic\">"+
                                                                              "    <one:PageSize>"+
                                                                              "      <one:Automatic />"+
                                                                              "    </one:PageSize>"+
                                                                              "    <one:RuleLines visible=\"false\" />"+
                                                                              "  </one:PageSettings>"+
                                                                              "  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\">"+
                                                                              "    <one:Position x=\"54.0\" y=\"122.400001525879\" z=\"0\" />"+
                                                                              "    <one:Size width=\"731.7882080078124\" height=\"35.37544250488281\" />"+
                                                                              "    <one:OEChildren>"+
                                                                              "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:14:07.000Z\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\" alignment=\"left\">"+
                                                                              "        <one:Table bordersVisible=\"true\" hasHeaderRow=\"true\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\">"+
                                                                              "          <one:Columns>"+
                                                                              "            <one:Column index=\"0\" width=\"445.8600769042969\" isLocked=\"true\" />"+
                                                                              "            <one:Column index=\"1\" width=\"277.7980651855469\" isLocked=\"true\" />"+
                                                                              "          </one:Columns>"+
                                                                              "          <one:Row lastModifiedTime=\"2014-07-05T03:15:19.000Z\">"+
                                                                              "            <one:Cell lastModifiedTime=\"2014-07-05T03:15:19.000Z\" lastModifiedByInitials=\"ID\">"+
                                                                              "              <one:OEChildren>"+
                                                                              "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:14:07.000Z\" lastModifiedTime=\"2014-07-05T03:15:19.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
                                                                              "                  <one:T><![CDATA[<span"+
                                                                              "style='font-weight:bold'>Task</span>]]></one:T>"+
                                                                              "                </one:OE>"+
                                                                              "              </one:OEChildren>"+
                                                                              "            </one:Cell>"+
                                                                              "            <one:Cell lastModifiedTime=\"2014-07-05T03:15:19.000Z\" lastModifiedByInitials=\"ID\">"+
                                                                              "              <one:OEChildren>"+
                                                                              "                <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:14:07.000Z\" lastModifiedTime=\"2014-07-05T03:15:19.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
                                                                              "                  <one:T><![CDATA[<span"+
                                                                              "style='font-weight:bold'>Date</span>]]></one:T>"+
                                                                              "                </one:OE>"+
                                                                              "              </one:OEChildren>"+
                                                                              "            </one:Cell>"+
                                                                              "          </one:Row>"+
                                                                              "        </one:Table>"+
                                                                              "      </one:OE>"+
                                                                              "      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-05T03:15:27.000Z\" lastModifiedTime=\"2014-07-05T03:15:27.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
                                                                              "        <one:T><![CDATA[]]></one:T>"+
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