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
        public void AugmentPage(OneNoteApp ona, XDocument pageContentInXml, OneNotePageCursor cursor)
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
                smartTodo.SetProcessed(ona);

                // TODO: Add TEST for missing page.
                var sourcePageContent = ona.GetPageContentAsXDocument(smartTodo.ParentPageId);

                // The cursor on the Get call indicates what location the smartags belong at. We're hacking here.
                var hackShouldBuildCursorForParentPage = cursor;
                var smartTagsOnSourcePage = SmartTag.Get(sourcePageContent, hackShouldBuildCursorForParentPage);

                // TODO: Add Test For missing smartTag
                var smartTag = smartTagsOnSourcePage.FirstOrDefault(st => st.ModelPageId == smartTodo.ParentModelId);
                if (smartTag == default(SmartTag))
                {
                    continue;
                }
                smartTag.SetCompleted(ona);
            }
        }
    }
}