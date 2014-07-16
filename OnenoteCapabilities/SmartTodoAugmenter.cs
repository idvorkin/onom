using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Xml.Serialization;
using Microsoft.Office.Interop.OneNote;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class SmartTodoAugmenter:IPageAugmenter
    {
        public static string CreateSmartTodoLink(OneNoteApp ona, SmartTag smartTag)
        {
            // should look like |__| where the first | encodes the important information and __ is a hyper link to the parent.
            var smartTodoTemplate = "<a href=http://smartTodo?{0}>|</a><a href={1}>____</a><a href=donotgohere>|</a>";
            var d = new Dictionary<string, string>();
            d["Processed"] = Boolean.FalseString;
            d["ParentPageId"] = smartTag.CursorLocation.PageId;

            // If I put this on the page, it will break smartTag searching -- need to figure that out later.
            // d["ParentModelId"] = smartTag.ModelPageId;

            // TODO Suck out section without doing getheirarchy in the future.
            var sectionAsXML = ona.GetHierarchy(smartTag.CursorLocation.SectionId, HierarchyScope.hsSelf);
            var section = OneNoteApp.XMLDeserialize<Section>(sectionAsXML);
            var pageAsXML = ona.GetHierarchy(smartTag.CursorLocation.PageId, HierarchyScope.hsSelf);
            var page = OneNoteApp.XMLDeserialize<Page>(pageAsXML);

            string linkToParentPage = OneNoteApp.OneNoteLinkToPage(page.name, section);
            var smartTodoXML = String.Format(smartTodoTemplate, ToQueryString(d), linkToParentPage);
            return smartTodoXML;
        }
        private static string ToQueryString(Dictionary<string,string> keyValuePairs)
        {
            var nameEqualValues = keyValuePairs.Select(x => x.Key + "=" + x.Value).ToArray();
            return "?" + string.Join("&", nameEqualValues);
        }
        public void AugmentPage(OneNoteApp ona, XDocument pageContentInXml, OneNotePageCursor cursor)
        {
            throw new NotImplementedException();
        }
    }
}