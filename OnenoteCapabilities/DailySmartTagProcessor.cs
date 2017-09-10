using System;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class DailySmartTagProcessor : ISmartTagProcessor
    {
        private SettingsDailyPages settings;
        private Section dailySection;

        public DailySmartTagProcessor(SettingsDailyPages settings)
        {
            this.settings = settings;
            this.dailySection = OneNoteApplication.Instance.GetNotebook(settings.DailyPagesNotebook)
                .PopulatedSection(settings.DailyPagesSection);

        }
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return (st.TagName() == "today");
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            var todayPageTitle = settings.DayPageTitleFromDate(DateTime.Now);
            var dailyPage = dailySection.GetPage(todayPageTitle);
            var todayPage = dailySection.GetPage(todayPageTitle);

            // IMPORTANT: The smartTag has a cache of the pageContent, so if we're writing the todo on the smartTag page via another pageContent, it will be overwritten.
            var dailyPageContent = (dailyPage.ID == cursor.PageId) ? pageContent : OneNoteApplication.Instance.GetPageContentAsXDocument(dailyPage);

            // HACK: Need to find the table of interest with a better method.
            // table 0 is the tasks table which is at the top of the page. If it moves down the table number changes - GROAN.
            var hackTableToAddTasksTo = 2;

            DumbTodo.AddToPageFromDateEnableSmartTag(dailyPageContent, smartTag, tableOnPage:hackTableToAddTasksTo);

            smartTag.SetLinkToPageId(todayPage.ID);
        }

        public string HelpLine()
        {
            return "<b>#today</b> add a SmartTodo for today from the rest of this line";
        }
    }
}
