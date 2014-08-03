using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteObjectModelTests
{
    class InkTagsTestPageContent:IPageContentAsText
    {
        public string firstLine()
        {
            return
                "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{0}\" name=\"{1}\" dateTime=\"2014-06-28T18:47:43.000Z\" lastModifiedTime=\"2014-07-02T13:10:47.000Z\" pageLevel=\"2\" isCurrentlyViewed=\"true\" lang=\"en-US\">";
        }

        public string restOfLines()
        {
            return "  <one:QuickStyleDef index=\"0\" name=\"PageTitle\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri Light\" fontSize=\"20.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
"  <one:QuickStyleDef index=\"1\" name=\"p\" fontColor=\"automatic\" highlightColor=\"automatic\" font=\"Calibri\" fontSize=\"11.0\" spaceBefore=\"0.0\" spaceAfter=\"0.0\" />"+
"  <one:PageSettings RTL=\"false\" color=\"automatic\">"+
"    <one:PageSize>"+
"      <one:Automatic />"+
"    </one:PageSize>"+
"    <one:RuleLines visible=\"false\" />"+
"  </one:PageSettings>"+
"  <one:Title lang=\"en-US\">"+
"    <one:OE author=\"Igor Dvorkin\" authorInitials=\"ID\" authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-25T06:36:27.000Z\" lastModifiedTime=\"2014-07-25T06:36:27.000Z\" alignment=\"left\" quickStyleIndex=\"0\">"+
"      <one:T><![CDATA[Scratch]]></one:T>"+
"    </one:OE>"+
"  </one:Title>"+
"  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-08-03T04:09:52.000Z\">"+
"    <one:Position x=\"54.0\" y=\"122.400001525879\" z=\"0\" />"+
"    <one:Size width=\"534.0751953125\" height=\"268.5542907714844\" isSetByUser=\"true\" />"+
"    <one:OEChildren>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:22:02.000Z\" lastModifiedTime=\"2014-07-28T23:22:02.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:List>"+
"          <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"        </one:List>"+
"        <one:T><![CDATA[People: ]]></one:T>"+
"        <one:OEChildren>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[EricaL]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[LiFenWu]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[SeanSe]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[IgorD]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[AmmonL]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[ToriS]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[SriRama]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[MaSudame]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[LarryS]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[DrRoebin]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[ZhiZhan]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[Devesh]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[KenMa ]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[JohnDoe]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[JaneDoe]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[ScotK]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[Descapa]]></one:T>"+
"          </one:OE>"+
"          <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:42.000Z\" lastModifiedTime=\"2014-07-28T23:21:42.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"            <one:List>"+
"              <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"            </one:List>"+
"            <one:T><![CDATA[Nmyre]]></one:T>"+
"          </one:OE>"+
"        </one:OEChildren>"+
"      </one:OE>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-07-28T23:21:57.000Z\" lastModifiedTime=\"2014-07-28T23:21:57.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:List>"+
"          <one:Bullet bullet=\"2\" fontSize=\"11.0\" />"+
"        </one:List>"+
"        <one:T><![CDATA[Next Setting]]></one:T>"+
"      </one:OE>"+
"    </one:OEChildren>"+
"  </one:Outline>"+
"  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-08-03T04:09:53.000Z\">"+
"    <one:Position x=\"348.3921203613281\" y=\"197.3905487060547\" z=\"2\" />"+
"    <one:Size width=\"72.0\" height=\"19.13385772705078\" isSetByUser=\"true\" />"+
"    <one:OEChildren>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-03T04:09:53.000Z\" lastModifiedTime=\"2014-08-03T04:09:53.000Z\" alignment=\"left\" quickStyleIndex=\"1\" style=\"font-family:Calibri;font-size:11.0pt\">"+
"        <one:InkWord recognizedText=\"I\" x=\"0.0\" y=\"0.0\" inkOriginX=\"-348.4063110351562\" inkOriginY=\"-197.4047241210937\" width=\"7.71023941040039\" height=\"19.13385772705078\">"+
"          <one:Data>APoCHQIWOAGAAVjPVIrml8VPjwb4utLhmyLeuxOL5V6zSJyboSVat1ES+oeVMvLsF0G4wBa/8yv0"+
"er/scczMIrFBiQ67EyzueqHEZm2nDbJYSZodTkA0BmwQRVOALHoY1USZZUo0BjFbMxSVQJ6SzJhL"+
"tWqQHBLFWlUqABofi0JwRYR1gL7QRdkEAwtIEES1u/AHRWRGZAUJOAtPZGVmZ2hrGQs4CQD+/wMA"+
"AAAAAE8JBGkQPoBHADAAGzQAAxEwAB3XAABKCjaSAuAPgDYAYgAAagMVRgBAawMVRgBACpwBHIP+"+
"TtT+TtXzt7G2ic6aaacY3vaF+JXL4ls8likk01mCxWMqmlwsMc8+Lltws9cloIftfuZ/537nf+h/"+
"z9wP+fs/s/Z+z8z9n7Pzlfs/cD/n7gf3hF/z94BfgAkECgARIKBLtsbQrs8BBRRGAAAAAAUURgSA"+
"AAASABEgWimGgdDoy0KT9DK90BFHRAo/QCLeigjTYGgVwAMVRgBA"+
"</one:Data>"+
"        </one:InkWord>"+
"      </one:OE>"+
"    </one:OEChildren>"+
"  </one:Outline>"+
"  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-08-03T04:09:53.000Z\">"+
"    <one:Position x=\"344.0716552734375\" y=\"221.1283416748047\" z=\"1\" />"+
"    <one:Size width=\"73.302734375\" height=\"13.42771339416504\" />"+
"    <one:OEChildren>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-03T04:09:53.000Z\" lastModifiedTime=\"2014-08-03T04:09:53.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:T><![CDATA[              ]]></one:T>"+
"      </one:OE>"+
"    </one:OEChildren>"+
"  </one:Outline>"+
"  <one:Outline author=\"Igor Dvorkin\" authorInitials=\"ID\" lastModifiedBy=\"Igor Dvorkin\" lastModifiedByInitials=\"ID\" lastModifiedTime=\"2014-08-03T04:10:17.000Z\">"+
"    <one:Position x=\"126.6944885253906\" y=\"533.0692749023437\" z=\"3\" />"+
"    <one:Size width=\"256.1386108398437\" height=\"17.54647064208984\" isSetByUser=\"true\" />"+
"    <one:OEChildren>"+
"      <one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-08-03T04:10:07.000Z\" lastModifiedTime=\"2014-08-03T04:10:17.000Z\" alignment=\"left\" quickStyleIndex=\"1\">"+
"        <one:InkWord recognizedText=\"#info\" style=\"font-family:Calibri;font-size:11.0pt\" x=\"0.0\" y=\"0.0\" inkOriginX=\"-126.7086639404297\" inkOriginY=\"-533.0834350585937\" width=\"71.94329833984374\" height=\"17.54647064208984\">"+
"          <one:Data>AMUNHQPEAT4BgAFYz1SK5pfFT48G+LrS4Zsi3rsTi+Ves0icm6ElWrdREvqHlTLy7BdBuMAWv/Mr"+
"9Hq/7HHMzCKxQYkOuxMs7nqhxGZtpw2yWEmaHU5ANAZsEEVTgCx6GNVEmWVKNAYxWzMUlUCeksyY"+
"S7VqkBwSxVpVKgAaH4tCcEWEdYC+0EXZBAMLSBBEtbvwB0VkRmQFCTgLT2RlZmdoaxkLOAkA/v8D"+
"AAAAAABPCQRpET9AIkXYAASXYAAG2AAAS+EAAEobNkendm3gEendmXAEend03gEendmygEend0AA"+
"AGoDFUYAQGsDFUYAQAqxASSE/CSl+EmPMsybd2e17XstnzWYYC629ntbfmCF+gtj6C+9Zi7Z0zCS"+
"W2YSq7BYSaSeC/B24WeWC+Su+W0Ah+9+53/lfuZ/5n/P3A/5+4n9n9yPzP+fs/5+Z+z/n7P+fs/O"+
"B+z9xvzgf3gl/z9wPwAJBAoAESDQvxjP0K7PAQUURgAAAAAFFEYEwAAAEgARIABqEsJZwr9CgnyF"+
"om9f5cIKABEgVl5iQD24bUEDFUYAQAqbARyE/CmZ+FKXd10cFy5a2e1mbTlsuuCF+gdr6B/YsYmk"+
"swk+ExWCwk088tuLwst8Ft8toIfnfuN/6X7jf+l/cT+z9n7gfs/M/Z/Z+z9n5xv2fs/Z+5394Tf8"+
"/AkECgARIGBGUs/Qrs8BBRRGAAAAAAUURgUAAAASABEgx7HuOEx7wk2nm/7U7jvxiwoAESAYv3JA"+
"y0tsQQMVRgBACo4BFIbwjVeEZumj5ygUktOwkLExtPEzsfLw4Ib0MRehjm9rpOHhoeFh4aEnImWg"+
"Y6EmoYCH8AP3O/8D9zP/G/s/OJ+Z/z8cD88Fv7vfs/AJBAoAESDgUoTP0K7PAQUURgAAAAAFFEYF"+
"QAAAEgARIGNcb3LGyUNHtUsHqM6DaiYKABEge71eQKeGbkEDFUYAQAqHARKH8JBXhH1sFPh0EhEO"+
"hkIi8BhEHiMAh8PAhP0Pifofb2ZbwvGFY1yUraW4h/Ab90v7pf+J/z9wP2fs/OR+z8Z/z94WfwAJ"+
"BAoAESCwXq3P0K7PAQUURgAAAAAFFEYFgAAAEgARIJ53TO0/EvZBph3YrKX9syUKABEgi8ViQNug"+
"c0EDFUYAQApqCoT8Mqn4ZP9rMGaE/QXF+gvfUMmXJqCH8Bv3S/5+6H/hfnQ/5+7H/PwJBAoAESCA"+
"XALQ0K7PAQUURgAAAAAFFEYGAAAAEgARIIK8G6GMaHhAqjC4EIM6uv0KABEgxvyHQMUVbkEDFUYA"+
"QAqLARaE/DLl+GXPdvzGa97EYF7Z73CE/Qil+hFOkNVN2ieLXGeesddbm0CH8CP3S/8L/0v/G/5+"+
"437P2fs/GfnO/Z+4H94Uf3gR+z8ACQQKABEgkALZz9CuzwEFFEYAAAAABRRGBcAAABIAESDWxQRm"+
"GfxES7QSXvUjWXB7CgARIIhdiEB1h3BBAxVGAEAKxQEog/4cNv4cJ/E8UAN723prtrWM9t63MbN6"+
"3oCH9CYHoRrpV/hEBg8Ag8Fg8HgsFg8Ng8bjcFh8OhkKi0YikUjEcg0KgsHhsFg8Ai8Ah/AL90P+"+
"f+l/4H/P3C/Z+437P3C/mfuB/HC/5+Z+4H9wP7hf8/Z/z9xvzifs/cL9n7jf8/eD38AJBAoAESDw"+
"1xvQ0K7PAQUURgAAAAAFFEYGQAAAEgARIH6Aet/i7ONNgCRswse167EKABEgKy+RQMb8b0EDFUYA"+
"QAp0DoT8QB34gH+ddTBmsRrchP0LOfoWH4thW9K11ZCH8BP3Q/8D90v+f2fnC/Z+6n90v+fgCQQK"+
"ABEg4ByH0NCuzwEFFEYAAAAABRRGBsAAABIAESCnKSlrthNPQJjPhGuYIiUOCj9AI+vaBOzjaC5A"+
"AxVGAEAKkAEahPxCCfiEVZr3tvzLsFLWjotbNalb54b0OTehyfAthAwsPNycbHwEVAQkZHQUdGQ0"+
"VDQ0VFiH8Dv3W/5+7H8DPzP7Pwz8435n7P3hd/z8CQQKABEgIFZZ0NCuzwEFFEYAAAAABRRGBoAA"+
"ABIAESBcQpHEJyJkR45Ll70cWMbkCj9AIyWCBPXe+C2gAxVGAEAKnAEahvEXR4i4bapoIaGiYCGh"+
"5GFj5ePh52hQEJDQ1BCAhfoa0+hnVqZZr7YL5Vtk8tUE1WEswk80IIfnfuN/6X7jf+R/cT/n7P+f"+
"jgf8/M/cD/n5n5n7P+fvBL+ACQQKABEgIOep0NCuzwEFFEYAAAAABRRGBwAAABIAESB33sM4swuO"+
"RIl0H/JZTfMUCgARIML6pkC8KnJBAxVGAEA=</one:Data>"+
"        </one:InkWord>"+
"        <one:InkWord width=\"28.48818397521973\" height=\"17.26299285888672\">"+
"          <one:Space />"+
"        </one:InkWord>"+
"        <one:InkWord recognizedText=\"Canada\" style=\"font-family:Calibri;font-size:11.0pt\" x=\"100.7433090209961\" y=\"0.0\" inkOriginX=\"-100.714958190918\" width=\"107.4047088623047\" height=\"17.54647064208984\">"+
"          <one:Data>AJALHQOiAjwBgAFYz1SK5pfFT48G+LrS4Zsi3rsTi+Ves0icm6ElWrdREvqHlTLy7BdBuMAWv/Mr"+
"9Hq/7HHMzCKxQYkOuxMs7nqhxGZtpw2yWEmaHU5ANAZsEEVTgCx6GNVEmWVKNAYxWzMUlUCeksyY"+
"S7VqkBwSxVpVKgAaH4tCcEWEdYC+0EXZBAMLSBBEtbvwB0VkRmQFCTgLT2RlZmdoaxkLOAkA/v8D"+
"AAAAAABPCQRpED6ARN4QAAAAAAHK4AACawAASiM3h4d2HJhAR8O7Dk3gFGPLuw49mUAo67sOTtlA"+
"KOu7DkyAAABqAxVGAEBrAxVGAEAKqAEchfg1a+DUNh5777aU+FhmwEmGw2Gx2GsxF0MuBlCG1Vuu"+
"TwJCQEVDQ0lGxkZKRkNFQ0PCzcbKwsPGw8MAh+p+5H/ofuN/6n/P3E/5+4H5xP+fuF+Z+Z+Z/z9w"+
"v2fuJ+c7/n7wU/5+CQQKABEg0NL00NCuzwEFFEYAAAAABRRGB0AAABIAESCW36oOk+WmQpJa4CTs"+
"wHXfCgARIIreyEBTqWxBAxVGAEAKtAEmhfg9e+D7WHbE0112GsovnhlvnvnJyaYumusonnpAhfDk"+
"eGRyWvSUQzxw2zy22zwTzyz4SS7EYSaSaGDCrZZAh/Ab90v7of8/uB+Z/z9n7P2f2f8/Z/z9wP7P"+
"+fuF/z9wv+fuB+z8cr8z9n7wq/vBj9n4CQQKABEgcJ4p0dCuzwEFFEYAAAAABRRGB4AAABIAESCw"+
"d+nPPXNjQrXhqrnSTjUECgARIPHe1UD4SHFBAxVGAEAKvAEqg/4Srv4Sl9dwmDeu+jY7NaYjs3uY"+
"22MAhuGC4VnBmG46CgIdGwEfDwqTiYuNlYVCRkZHSUlHSUZDQiGlYWNh4eAhwIfwC/dD/wP/O/8L"+
"+5H7P2f3A/OB+HA/uB/Z+z9yP7P+fs/5+4n/P3A/Z+4X5yPzjfngt/cj9n4JBAoAESBAMmbR0K7P"+
"AQUURgAAAAAFFEYHwAAAEgARIHaFkfqzw85PuvFgjU6+WrcKABEg8xLjQIbcb0EDFUYAQAq8ASqE"+
"/DB1+GEPbS+e9q5M2a27NSC698Nblr2GbNa1IXRAhPhWPhUc9s2LVglOU898ccd22uW+PDglmtq1"+
"bNmSiHFvjz5gh+9+53/kfuZ/53/PzP3A/5+4X5wP2fs/OB/z9wv7P+fuF/z8cL+z+4X7PzP2fs/c"+
"j8Z+8Mv+fgkECgARIDAUo9HQrs8BBRRGAAAAAAUURggAAAASABEgxbCtTFma4USL0Bmbc5ptnQo/"+
"QCODBgeG2oguAAMVRgBACuEBMIbw3eeG+2XroCBh4aGho6WnoyQhIhEwszGx8vLx8vCwJDxEDLR0"+
"tLS05MR8JIQkBIyohuD84O+ajYaSlI2Sk4eNjYWNh5WNi52NRsZGws7CykVST1VVSE4lJKMjaOTl"+
"ZOVlYUCH8Bv3S/5+6X/if2f8/c78z+z/n7gf2f3A/jPxn7hf8/cL/n7P+fuF/cb8Z+5H7P3E/5+4"+
"X5n54Yf8/AkECgARIJB449HQrs8BBRRGAAAAAAUURghAAAASABEgLsRmoBmpL0KnI1WCtOdj1woA"+
"ESDBE/1AxchqQQMVRgBACtQBLoX4hSPiFxxusTXTWUYa6SW+mW+ee+/Hk8Ek12Ow2EguwlE1OAht"+
"hwt9IIbgjOBHor+UhEKjY+LlZWZiZOAj0pDpqGlJKMmpqijIeZj4Wdg4mLj4uBjYaNCH8Bv3S/8z"+
"/0P/K/Z+5H7P3E/uB/M/cL/n7P7Pxn7hf8/Z/z9yv7mfmf8/cr843/Pwcz88FP+fuF/ACQQKABEg"+
"gI8s0tCuzwEFFEYAAAAABRRGCIAAABIAESA3zXGMX0kWRYStfTleC5d6Cj9AI/ciCEQuyC3AAxVG"+
"AEA=</one:Data>"+
"        </one:InkWord>"+
"      </one:OE>"+
"    </one:OEChildren>"+
"  </one:Outline>"+
"</one:Page>";

        }
    }
}
