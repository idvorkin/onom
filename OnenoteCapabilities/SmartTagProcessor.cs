using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public interface ISmartTagProcessor
    {
        // TBD:Ensure matching functions don't overlap (easy source of bugs)
        bool ShouldProcess(SmartTag st);
        void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter);
    }

    public class PeopleSmartTagProcessor : ISmartTagProcessor
    {
        private SettingsPeoplePages settings;
        private Section peopleSection;
        private DumbTodo _dumbTodo;

        public PeopleSmartTagProcessor(OneNoteApp ona, SettingsPeoplePages settings)
        {
            this.settings = settings;
            _dumbTodo = new DumbTodo();
            this.peopleSection = ona.GetNotebooks()
                .Notebook.First(n => n.name == settings.PeoplePagesNotebook)
                .PopulatedSections(ona).First(s => s.name == settings.PeoplePagesSection);
        }

        public bool ShouldProcess(SmartTag st)
        {
            if (settings.People().Contains(personFromPersonTag(st))) return true;
            return false;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // TODO: Create Person Page if not Exists.
            var personPageTitle = settings.PersonNextTitle(personFromPersonTag(smartTag));

            // get PersonPage 
            var peoplePage = peopleSection.Page.First(p => p.name == personPageTitle);
            var newPageContent = "";
            smartTagAugmenter.ona.OneNoteApplication.GetPageContent(peoplePage.ID,out newPageContent);

            const int toPersonTableCountOnPage = 0;
            const int fromPersonTableCountOnPage = 1;

            var tableOnPage = IsFromPerson(smartTag) ? fromPersonTableCountOnPage : toPersonTableCountOnPage;
            _dumbTodo.AddDumbTodoToPage(smartTagAugmenter.ona,XDocument.Parse(newPageContent),smartTag.TextAfterTag(),tableOnPage:tableOnPage);
            smartTagAugmenter.AddLinkToSmartTag(smartTag, pageContent, peopleSection, personPageTitle);
        }

        public bool IsFromPerson(SmartTag smartTag)
        {
            return smartTag.TagName().StartsWith("from");
        }

        public string personFromPersonTag (SmartTag smartTag)
        {
            if (IsFromPerson(smartTag))
            {
                return smartTag.TagName().Substring(4);
            }
            else
            {
                return smartTag.TagName();
            }
        }
    }

    
    public class TopicSmartTagTopicProcessor: ISmartTagProcessor
    {
        private OneNoteApp ona;
        private TemplatePageCreator templatePageCreater;
        private readonly DumbTodo _dumbTodo;

        // TODO move from settingDailyPages to SettingTopicPages
        public TopicSmartTagTopicProcessor(OneNoteApp ona, SettingsDailyPages settings)
        {
            this.ona = ona;
            this.templatePageCreater = new TemplatePageCreator(ona, settings.TemplateNotebook, settings.TemplateSection, settings.DailyPagesNotebook, settings.DailyPagesSection);
            _dumbTodo = new DumbTodo();
        }
        public bool ShouldProcess(SmartTag st)
        {
            // Topic Processor should be added last because it will always make a smarttag.
            return true;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // Create Topic Page of name 
            var topicPageName = smartTag.TagName();
            var page = templatePageCreater.CreatePageIfNotExists(topicPageName,"Topic",1);
            var newPageContent = "";
            ona.OneNoteApplication.GetPageContent(page.ID,out newPageContent);
            _dumbTodo.AddDumbTodoToPage(ona,XDocument.Parse(newPageContent),smartTag.TextAfterTag());
            smartTagAugmenter.AddLinkToSmartTag(smartTag,pageContent,templatePageCreater.SectionForPages(), topicPageName);
        }
    }

    public class TwitterSmartTagProcessor: ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st)
        {
            return st.TagName().ToLower() == "tweet";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // do nothing.
        }
    }
}
