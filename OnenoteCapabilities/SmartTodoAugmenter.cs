using System;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Xml.Serialization;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class SmartTodoAugmenter:IPageAugmenter
    {
        public void AugmentPage(XDocument pageContentInXml, OneNotePageCursor cursor)
        {
            var smartTodos = SmartTodo.Get(pageContentInXml);

            // Only process the first time they are marked completed.
            var smartTodosToProcess =  smartTodos.Where(st=>st.IsProcessed == false && st.IsComplete).ToList();

            if (!smartTodosToProcess.Any())
            {
                return;
            }

            foreach (var smartTodo in smartTodosToProcess)
            {
                smartTodo.SetProcessed();

                XDocument sourcePageContent;
                try
                {
                    sourcePageContent = OneNoteApplication.Instance.GetPageContentAsXDocument(smartTodo.ParentPageId);
                }
                catch (COMException)
                {
                    // TODO: Add TEST for missing page.
                    // we throw if page is not found - TBD - test for exact error code.
                    continue;
                }

                // The cursor on the Get call indicates what location the smartags belong at. We're hacking here.
                var hackShouldBuildCursorForParentPage = cursor;
                var smartTagsOnSourcePage = SmartTag.Get(sourcePageContent, hackShouldBuildCursorForParentPage);

                // TODO: Add Test For missing smartTag
                var smartTag = smartTagsOnSourcePage.FirstOrDefault(st => st.ModelPageId == smartTodo.ParentModelId);
                if (smartTag == default(SmartTag))
                {
                    continue;
                }
                smartTag.SetCompleted();


                var page = OneNoteApplication.XMLDeserialize<Page>(pageContentInXml.ToString());
                var pageLink = OneNoteApplication.Instance.GetHyperLinkToObject(page.ID);
                var pageName = page.name;

                var creationText = String.Format("Completed smartTag  '{0} {3}' via smartTodo on page <a href='{1}'> {2} </a>", 
                        smartTag.TagName(), pageLink, pageName, smartTag.TextAfterTag());

                smartTag.AddEntryToModelPage(creationText);

            }
        }
    }
}
