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
        public SmartTagAugmenter(SettingsSmartTags settings,IEnumerable<ISmartTagProcessor> smartTagProcessors)
        {
            Debug.Assert(smartTagProcessors != null);

            this.smartTagProcessors = smartTagProcessors;
            this.settings = settings;

            var smartTagNotebook = OneNoteApplication.Instance.GetNotebook(settings.SmartTagStorageNotebook);
            this.smartTagModelSection = smartTagNotebook.PopulatedSection(settings.SmartTagModelSection);

            var templateNoteBook = OneNoteApplication.Instance.GetNotebook(settings.TemplateNotebook);
            this.smartTagTemplatePage = templateNoteBook.PopulatedSection(settings.TemplateSection).GetPage(this.settings.SmartTagTemplateName);
        }

        public void AugmentPage(XDocument pageContent, OneNotePageCursor cursor)
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
            var modelPage = OneNoteApplication.Instance.ClonePage(smartTagModelSection,smartTagTemplatePage , newModelPageName );
            // put a hyper-link to the page on the '#'
            smartTag.SetId(newModelPageName, modelPage.ID, smartTagModelSection);

            // TBD: get page attributes without the slow XMLDeserialize.
            var page = OneNoteApplication.XMLDeserialize<Page>(pageContent.ToString());
            var pageLink = OneNoteApplication.Instance.GetHyperLinkToObject(page.ID);
            var pageName = page.name;

            var creationText = String.Format("Instantiated model from tag '{0}' on page <a href='{1}'> {2} </a> with tag text:{3}", 
                    smartTag.TagName(), pageLink, pageName, smartTag.TextAfterTag());

            smartTag.AddEntryToModelPage(creationText);
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
        private Page smartTagTemplatePage;
    }
}
