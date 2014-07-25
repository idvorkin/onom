using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml.Linq;
using System.Xml.Serialization;
using Microsoft.Office.Interop.OneNote;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class SmartTodoAugmenter:IPageAugmenter
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
                throw new NotImplementedException();
            }
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

        // HACK: Should have two functions, 1 testing for completion, and another parsing the TODO link.
        //      In the MVP these will be merged.
        public static SmartTodo CreateFromXmlElement(XElement e)
        {
            // Hack - this understands how create SmartTodo works - add to helper function.
            var matches = Regex.Match(e.Value, "http://smarttodo?(.*)\\>.<a/>");
            var queryString = matches.Groups[0].Value;
            var smartTodoParams =  HttpUtility.ParseQueryString(queryString);

            return new SmartTodo()
            {
                Processed = bool.Parse(smartTodoParams["Processed"]),
                ParentPageId =  smartTodoParams["ParentPageId"],
                ParentModelId =  smartTodoParams["ParentModelId"],
                IsCompleted = true, // TODO: Parse is completed out properly.
            };
        }

        private static string ToQueryString(Dictionary<string,string> keyValuePairs)
        {
            var nameEqualValues = keyValuePairs.Select(x => x.Key + "=" + x.Value).ToArray();
            return "?" + string.Join("&", nameEqualValues);
        }
        public void AugmentPage(OneNoteApp ona, XDocument pageContentInXml, OneNotePageCursor cursor)
        {
            var smartTodoElements = pageContentInXml.DescendantNodes()
                .OfType<XElement>()
                .Where(r => r.Name.LocalName == "T" && r.Value.Contains("http://smartTodo"));

            var smartTodos = smartTodoElements.Select(CreateFromXmlElement);

            var smartTodosToProcess =  smartTodos.Where(st=>st.Processed == false && st.IsCompleted).ToList();

            if (!smartTodosToProcess.Any())
            {
                return;
            }

            foreach (var smartTodo in smartTodosToProcess)
            {
                smartTodo.SetProcessed(ona);
                // TODO: Add TEST for missing page.
                var sourcePageContent = ona.GetPageContentAsXDocument(smartTodo.ParentPageId);
                var smartTagsOnSourcePage = SmartTag.Get(sourcePageContent, cursor);

                // TODO: Add Test For missing smartTag
                var smartTag = smartTagsOnSourcePage.FirstOrDefault(st => st.ModelPageId == smartTodo.ParentModelId);
                if (smartTag == default(SmartTag))
                {
                    continue;
                }
                smartTag.SetCompleted(ona);
            }

            // smartTodo was processed, update the processed flags.
            ona.OneNoteApplication.UpdatePageContent(pageContentInXml.ToString());
        }

        private void MarkSmartTagAsCompleteInPageContent(SmartTag smartTag, XDocument sourcePageContent)
        {
            throw new NotImplementedException();
        }

        private void MarkSmartTodoAsProcessedInPageContent(SmartTodo smartTodo, XDocument pageContentInXml)
        {
            throw new NotImplementedException();
        }

    }
}