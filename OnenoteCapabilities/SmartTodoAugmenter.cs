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
    }
}