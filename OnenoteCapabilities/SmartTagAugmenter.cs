using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Xml;
using System.Xml.Linq;
using OnenoteCapabilities;
using OneNoteObjectModel;
using Page = OneNoteObjectModel.Page;

namespace OnenoteCapabilities
{
    public class SmartTagAugmenter:IPageAugmenter
    {
        public SmartTagAugmenter(OneNoteApp ona, SettingsSmartTags settings,IEnumerable<ISmartTagProcessor> smartTagProcessors)
        {
            Debug.Assert(smartTagProcessors != null);

            this.smartTagProcessors = smartTagProcessors;
            this.settings = settings;
            this.ona = ona;

            var smartTagNotebook = ona.GetNotebook(settings.SmartTagStorageNotebook);
            this.smartTagModelSection = smartTagNotebook.PopulatedSection(ona,settings.SmartTagModelSection);

            var templateNoteBook = ona.GetNotebook(settings.TemplateNotebook);
            this.smartTagTemplatePage = templateNoteBook.PopulatedSection(ona, settings.TemplateSection).GetPage(ona, this.settings.SmartTagTemplateName);
        }

        public void AugmentPage(OneNoteApp ona, XDocument pageContent, OneNotePageCursor cursor)
        {
            var smartTags = SmartTag.Get(pageContent, cursor);
            foreach (var smartTag in smartTags.Where(st=>!st.IsProcessed()))
            {
                AddToModel(smartTag, pageContent);
                ProcessSmartTag(smartTag, pageContent, cursor);
            }
        }

        public void AddToModel(SmartTag smartTag, XDocument pageContent)
        {
            // create a new page to represent the smart tag. 
            var newModelPageName = string.Format("Model: {0}", Guid.NewGuid());
            var modelPage = ona.ClonePage(smartTagModelSection,smartTagTemplatePage , newModelPageName );

            // put a hyper-link to the page on the '#'
            smartTag.SetId(ona,newModelPageName, modelPage.ID, smartTagModelSection);
                
            // TODO Refactor to somewhere else.
            // record human readable text on the Model Page.
            var modelPageContent = ona.GetPageContentAsXDocument(modelPage);

            // NOTE: This is a big perf hit - figure out how to refactor.
            var page = OneNoteApp.XMLDeserialize<Page>(pageContent.ToString());

            // TODO: Make this hyper-link.
            var creationText = String.Format("Instantiated model from tag '{1}' on page '{0}' with tag text:{2}", page.name, smartTag.TagName(),smartTag.TextAfterTag());
            DumbTodo.AddToPage(ona,modelPageContent, creationText, DateTime.Now);
        }

        /// <summary>
        /// The 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<string> GetAllHelpLines()
        {
            return smartTagProcessors.Select(p => p.HelpLine());
        }


        private void ProcessSmartTag(SmartTag smartTag, XDocument pageContent, OneNotePageCursor cursor)
        {
            foreach (var tagProcessor in smartTagProcessors)
            {
                if (tagProcessor.ShouldProcess(smartTag, cursor))
                {
                    tagProcessor.Process(smartTag,pageContent, this, cursor);
                    break;
                }
            }
        }

        public Section smartTagModelSection;
        public IEnumerable<ISmartTagProcessor> smartTagProcessors;
        private SettingsSmartTags settings;
        public OneNoteApp ona;
        private Page smartTagTemplatePage;

    }
}
