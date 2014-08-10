using System;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
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

        public string HelpLine()
        {
            return "<b>#today</b> add a SmartTodo for today from the rest of this line";
        }
    }
}