using System;
using System.Linq;
using System.Xml.Linq;

namespace OnenoteCapabilities
{
    public class ActionItemSmartTagProcessor : ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return st.TagName() == "ai";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            var tableOnPageToWriteTo = FindDumbTodoTableOnPage(pageContent);
            if (!tableOnPageToWriteTo.HasValue )
            {
                smartTag.AddContentAfter("this page does not support action items (can not find a table to use).");
                return;
            }

            DumbTodo.AddToPageFromDateEnableSmartTag(pageContent, smartTag,tableOnPageToWriteTo.Value);
            smartTag.SetLinkToPageId(cursor.PageId);
        }

        private int? FindDumbTodoTableOnPage(XDocument pageContentAsXml)
        {
            // For now, only work with one table on the  page.
            var tables = pageContentAsXml.DescendantNodes().OfType<XElement>().Where(e => e.Name.LocalName == "Table");
            if (tables.Count() == 1)
            {
                // This is brittle, but good enough for now. 
                // In the future figure out how to detect a smartTodo enabled table.
                return 0;
            }
            return null;
        }

        public string HelpLine()
        {
            return "<b>#ai</b> create an action item from the rest of the line.";
        }
    }
}
