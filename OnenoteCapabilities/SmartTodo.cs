using System;
using System.Collections.Generic;
using System.Diagnostics.Tracing;
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
        public bool IsProcessed;
        public string ParentPageId;
        public string ParentModelId;
        public bool IsComplete;
        public XElement Element;
        public XDocument PageContent;

        public void SetProcessed(OneNoteApp ona)
        {
            this.Element.Value = this.Element.Value.Replace("IsProcessed=False", "IsProcessed=True");
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
            smartTodoProperties["IsProcessed"] = Boolean.FalseString;
            smartTodoProperties["ParentPageId"] = smartTag.CursorLocation.PageId;
            smartTodoProperties["ParentModelId"] = smartTag.ModelPageId;
            smartTodoProperties["Trailer"] = "nonce";

            var sectionContaingSmartTag = OneNoteApp.XMLDeserialize<Section>(ona.GetHierarchy(smartTag.CursorLocation.SectionId, HierarchyScope.hsSelf));
            var pageContainingSmartTag = OneNoteApp.XMLDeserialize<Page>(ona.GetHierarchy(smartTag.CursorLocation.PageId, HierarchyScope.hsSelf));

            // should look like <--. where the . encodes information and <-- is a hyper link to the parent.
            string linkToParentPage = OneNoteApp.OneNoteLinkToPage(pageContainingSmartTag.name, sectionContaingSmartTag);
            var smartTodoTemplate = "<a href={1}>&lt;---</a><a href=http://smartTodo{0}>.</a> ";
            var smartTodoXML = String.Format(smartTodoTemplate, ToQueryString(smartTodoProperties), linkToParentPage);
            return smartTodoXML;
        }
        public static bool IsSmartTodo(XElement smartTodoElement)
        {
            if (smartTodoElement.Name.LocalName != "T")
            {
                return false;
            }
            var matches = Regex.Match(smartTodoElement.Value,  SmartTodoRegexpParamCapture);
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

        public static readonly string SmartTodoRegexpParamCapture = "http://smartTodo\\?(.*)>\\.</a>";
            

        public static SmartTodo FromXmlElement(XElement smartTodoElement, XDocument pageContent)
        {
            // Hack - this understands how create SmartTodo works - add to helper function.
            var matches = Regex.Match(smartTodoElement.Value, SmartTodoRegexpParamCapture);
            var queryString = matches.Groups[1].Value;
            var decodedQueryString = HttpUtility.HtmlDecode(HttpUtility.UrlDecode(queryString));
            var smartTodoParams = HttpUtility.ParseQueryString(decodedQueryString);

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
                IsProcessed = Boolean.Parse(smartTodoParams["IsProcessed"]),
                ParentPageId = smartTodoParams["ParentPageId"],
                ParentModelId = smartTodoParams["ParentModelId"],
                IsComplete = isCompleted,
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