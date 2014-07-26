using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public struct SmartTodo
    {
        public bool Processed;
        public string ParentPageId;
        public string ParentModelId;
        public bool IsCompleted;
        public XElement Element;
        public XDocument PageContent;


        public void SetProcessed(OneNoteApp ona)
        {
            this.Element.Value.Replace("Processed=False", "Processed=True");
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }

        public static IEnumerable<SmartTodo> Get(XDocument pageContentInXml)
        {
            var smartTodoElements = pageContentInXml.DescendantNodes().OfType<XElement>().Where(IsSmartTodo);
            return smartTodoElements.Select(e=>FromXmlElement(e,pageContentInXml));
        }

        public static string CreateSmartTodoLink(OneNoteApp ona, SmartTag smartTag)
        {
            var smartTodoProperties = new Dictionary<string, string>();
            smartTodoProperties["Processed"] = Boolean.FalseString;
            smartTodoProperties["ParentPageId"] = smartTag.CursorLocation.PageId;
            smartTodoProperties["ParentModelId"] = smartTag.ModelPageId;

            var sectionContaingSmartTag = OneNoteApp.XMLDeserialize<Section>(ona.GetHierarchy(smartTag.CursorLocation.SectionId, HierarchyScope.hsSelf));
            var pageContainingSmartTag = OneNoteApp.XMLDeserialize<Page>(ona.GetHierarchy(smartTag.CursorLocation.PageId, HierarchyScope.hsSelf));

            // should look like <--. where the . encodes information and <-- is a hyper link to the parent.
            string linkToParentPage = OneNoteApp.OneNoteLinkToPage(pageContainingSmartTag.name, sectionContaingSmartTag);
            var smartTodoTemplate = "<a href={1}>&lt;---</a><a href=http://smartTodo?{0}>.</a> ";
            var smartTodoXML = String.Format(smartTodoTemplate, ToQueryString(smartTodoProperties), linkToParentPage);
            return smartTodoXML;
        }
        public static bool IsSmartTodo(XElement smartTodoElement)
        {
            var matches = Regex.Match(smartTodoElement.Value, "http://smarttodo?(.*)\\>.<a/>");
            if (!matches.Success)
            {
                return false;
            }
            
            var todoBox = smartTodoElement.Parent.DescendantNodes().OfType<XElement>().FirstOrDefault(e => e.Name.LocalName == "Tag");

            if (todoBox == default(XElement))
            {
                return false;
            }

            return true;
        }
            

        public static SmartTodo FromXmlElement(XElement smartTodoElement, XDocument pageContent)
        {
            // Hack - this understands how create SmartTodo works - add to helper function.
            var matches = Regex.Match(smartTodoElement.Value, "http://smarttodo?(.*)\\>.<a/>");
            var queryString = matches.Groups[0].Value;
            var smartTodoParams = HttpUtility.ParseQueryString(queryString);

            // SmartTags in onenote look like
            // OE
            //  +> Tag
            //  +> T

            var todoBox = smartTodoElement.Parent.DescendantNodes().OfType<XElement>().FirstOrDefault(e => e.Name.LocalName == "Tag");

            if (todoBox == default(XElement))
            {
                throw new InvalidDataException("Not a smartTodo");
            }

            var isCompleted = (todoBox.Attribute("completed").Value == "true");

            return new SmartTodo()
            {
                Processed = Boolean.Parse(smartTodoParams["Processed"]),
                ParentPageId = smartTodoParams["ParentPageId"],
                ParentModelId = smartTodoParams["ParentModelId"],
                IsCompleted = isCompleted,
                Element = smartTodoElement,
                PageContent = pageContent,
            };
        }
        private static string ToQueryString(Dictionary<string,string> keyValuePairs)
        {
            var nameEqualValues = keyValuePairs.Select(x => x.Key + "=" + x.Value).ToArray();
            return "?" + String.Join("&", nameEqualValues);
        }
    }
}