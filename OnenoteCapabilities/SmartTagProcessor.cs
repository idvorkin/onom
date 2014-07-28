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
        bool ShouldProcess(SmartTag st, OneNotePageCursor cursor);
        void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor);
    }

    public class DailySmartTagProcessor : ISmartTagProcessor
    {
        private SettingsDailyPages settings;
        private OneNoteApp ona;
        private Section dailySection;

        public DailySmartTagProcessor(OneNoteApp ona, SettingsDailyPages settings)
        {
            this.ona = ona;
            this.settings = settings;
            this.dailySection = ona.GetNotebook(settings.DailyPagesNotebook)
                .PopulatedSection(ona, settings.DailyPagesSection);

        }
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return (st.TagName() == "today");
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            var todayPageTitle = settings.DayPageTitleFromDate(DateTime.Now);
            var dailyPage = dailySection.GetPage(ona, todayPageTitle);

            // IMPORTANT: The smartTag has a cache of the pageContent, so if we're writing the todo on the smartTag page via another pageContent, it will be overwritten.
            var dailyPageContent = (dailyPage.ID == cursor.PageId) ? pageContent : ona.GetPageContentAsXDocument(dailyPage);

            // HACK: Need to find the table of interest with a better method.
            // table 0 is the tasks table which is at the top of the page. If it moves down the table number changes - GROAN.
            var hackTableToAddTasksTo = 0;

            DumbTodo.AddToPageFromDateEnableSmartTag(smartTagAugmenter.ona, dailyPageContent, smartTag, tableOnPage:hackTableToAddTasksTo);

            smartTag.SetLink(ona,dailySection,todayPageTitle);
        }
    }

    // NOTE: I'm hardcoding this today to get something working experimently. I'll hack on it once I make some progress.

    public class PeopleSmartTagProcessor : ISmartTagProcessor
    {
        private SettingsPeoplePages settings;
        private Section peopleSection;
        private OneNoteApp ona;

        public PeopleSmartTagProcessor(OneNoteApp ona, SettingsPeoplePages settings)
        {
            this.ona = ona;
            this.settings = settings;
            this.peopleSection = ona.GetNotebook(settings.PeoplePagesNotebook)
                .PopulatedSection(ona, settings.PeoplePagesSection);
        }

        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            if (settings.People(ona).Contains(personFromPersonTag(st))) return true;
            return false;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            // TODO: Create Person Page if not Exists.
            var personPageTitle = settings.PersonNextTitle(personFromPersonTag(smartTag));

            // get PersonPage 

            var peoplePage = peopleSection.GetPage(ona,personPageTitle);
            // IMPORTANT: The smartTag has a cache of the pageContent, so if we're writing the todo on the smartTag page via another pageContent, it will be overwritten.
            var peoplePageContent = (peoplePage.ID == cursor.PageId) ? pageContent : ona.GetPageContentAsXDocument(peoplePage);

            const int toPersonTableCountOnPage = 0;
            const int fromPersonTableCountOnPage = 1;

            var tableOnPage = IsFromPerson(smartTag) ? fromPersonTableCountOnPage : toPersonTableCountOnPage;
            DumbTodo.AddToPageFromDateEnableSmartTag(ona, peoplePageContent, smartTag, tableOnPage: tableOnPage);


            smartTag.SetLink(ona, peopleSection, personPageTitle);
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
        private SettingsTopicPages settings;

        // TODO move from settingDailyPages to SettingTopicPages
        public TopicSmartTagTopicProcessor(OneNoteApp ona, SettingsTopicPages settings)
        {
            this.ona = ona;
            this.templatePageCreater = new TemplatePageCreator(ona, settings.TemplateNotebook, settings.TemplateSection, settings.TopicNotebook, settings.TopicSection);
            this.settings = settings;
        }
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            // Topic Processor should be added last because it will always make a smarttag.
            return true;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            // Create Topic Page of name 
            var topicPageName = smartTag.TagName();
            var topicPage = templatePageCreater.CreatePageIfNotExists(topicPageName,settings.TopicTemplateName,1);
            var topicPageContent = ona.GetPageContentAsXDocument(topicPage);

            DumbTodo.AddToPageFromDateEnableSmartTag(ona, topicPageContent,smartTag);

            if (OneNoteApp.IsSamePage (pageContent, topicPageContent))
            {
                // HACK: When the smartTag is on the current dailyPage, we put the TODO on the peoplePageContent, but then we put the link on pageContent, which doesn't have the changes to PageContent.
                pageContent = topicPageContent;
            }

            smartTag.SetLink(ona,templatePageCreater.SectionForPages(), topicPageName);
        }
    }
}
