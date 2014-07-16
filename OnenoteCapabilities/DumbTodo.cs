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

        public static void AddToPage(OneNoteApp ona, XDocument pageContentAsXML, string todo, DateTime? dueDate=null, int tableOnPage=0)
        {
            AddTodoTagToPageIfRequired(ona, pageContentAsXML);
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

        // If the page does not contain a todo tag, writing a todo will fail, so we need to add it explicitly if it does not exist.
        private static void AddTodoTagToPageIfRequired(OneNoteApp ona, XDocument pageContentAsXml)
        {
            var tagDefPresent = pageContentAsXml.DescendantNodes().OfType<XElement>().Where(e => e.Name.LocalName == "TagDef").Any();
            if (tagDefPresent)
            {
                return;
            }
            var todoTagDef = "  <one:TagDef index=\"0\" type=\"0\" symbol=\"3\" fontColor=\"automatic\" highlightColor=\"none\" name=\"To Do\" xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"/>";
            var tagDefAsXml = XDocument.Parse(todoTagDef);
            pageContentAsXml.DescendantNodes().OfType<XElement>().First(e => e.Name.LocalName == "QuickStyleDef").AddBeforeSelf(tagDefAsXml.Root);
        }

        public static void AddToPageFromDateEnableSmartTag(OneNoteApp ona , XDocument pageContent, SmartTag smartTag, int tableOnPage=0)
        {
            var parsedDate = HumanizedDateParser.ParseDateAtEndOfSentance(smartTag.TextAfterTag());

            // HACK - TBD Design this.
            var smartTodoPrefix = SmartTodoAugmenter.CreateSmartTodoLink(ona, smartTag);

            if (parsedDate.Parsed)
            {
                DumbTodo.AddToPage(ona, pageContent, smartTodoPrefix+parsedDate.SentanceWithoutDate, parsedDate.date, tableOnPage);
            }
            else
            {
                DumbTodo.AddToPage(ona, pageContent, smartTodoPrefix+smartTag.TextAfterTag(), tableOnPage:tableOnPage);
            }
        }
    }
}