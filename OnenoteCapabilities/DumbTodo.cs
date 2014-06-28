using System;
using System.Linq;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class DumbTodo
    {
        public DumbTodo()
        {
        }

        public void AddDumbTodoToPage(OneNoteApp ona, XDocument pageContentAsXML, string todo, DateTime? dueDate=null, int tableOnPage=0)
        {
            var rowTemplate = "<one:Row lastModifiedTime=\"2014-06-28T06:11:19.000Z\" xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"> " +
                              "<one:Cell lastModifiedTime=\"2014-06-28T06:11:19.000Z\"  lastModifiedByInitials=\"ID\"> " +
                              "<one:OEChildren> " +
                              "<one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:11:19.000Z\" lastModifiedTime=\"2014-06-28T06:11:19.000Z\" alignment=\"left\">" +
                              "<one:Tag index=\"0\" completed=\"{2}\" disabled=\"false\" creationDate=\"2014-06-28T06:11:25.000Z\" completionDate=\"2014-06-28T06:27:01.000Z\" />"+
                              "<one:T><![CDATA[{0}]]></one:T> " +
                              "</one:OE> " +
                              "</one:OEChildren> " +
                              "</one:Cell> " +
                              "<one:Cell lastModifiedTime=\"2014-06-28T06:11:13.000Z\" lastModifiedByInitials=\"ID\"> " +
                              "<one:OEChildren> " +
                              "<one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:11:13.000Z\" lastModifiedTime=\"2014-06-28T06:11:13.000Z\" alignment=\"left\">" +
                              "<one:T><![CDATA[{1}]]></one:T> " +
                              "</one:OE> " +
                              "</one:OEChildren> " +
                              "</one:Cell> " +
                              "</one:Row>";

            bool completed=false;

            var row = string.Format(rowTemplate,todo,dueDate != null ?  dueDate.Value.ToShortDateString(): "", completed.ToString().ToLower());
            var rowAsXML = XDocument.Parse(row);

            // Skip tables in DOM.
            var tableElement = pageContentAsXML.DescendantNodes() .OfType<XElement>() .Where(e => e.Name.LocalName == "Table").Skip(tableOnPage);

            // Add row after the first row (which is assumed to be a header)
            tableElement.DescendantNodes().OfType<XElement>().First(e => e.Name.LocalName=="Row").AddAfterSelf(rowAsXML.Root);
            ona.OneNoteApplication.UpdatePageContent(pageContentAsXML.ToString());
        }
    }
}