using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class DailyPages
    {
        private SettingsDailyPages _settingsDailyPages;

        public DailyPages(OneNoteApp ona, SettingsDailyPages settingsDailyPages)
        {
            this._settingsDailyPages = settingsDailyPages;
            _templatePageCreator = new TemplatePageCreator(ona,
                    templateNotebook:settingsDailyPages.TemplateNotebook, 
                    templateSection:settingsDailyPages.TemplateSection, 
                    pagesNotebook:settingsDailyPages.DailyPagesNotebook, 
                    pagesSection:settingsDailyPages.DailyPagesSection);
        }

        private readonly TemplatePageCreator _templatePageCreator;

        public void GotoTodayPage()
        {
            string todayPageTitle = _settingsDailyPages.TodayPageTitle();
            _templatePageCreator.GotoOrCreatePage(todayPageTitle, _settingsDailyPages.TemplateDailyPageTitle, 2);
        }

        public void GotoYesterday()
        {
            string yesterdayPageTitle = _settingsDailyPages.DayPageTitleFromDate(DateTime.Now - TimeSpan.FromDays(1.0));
            _templatePageCreator.GotoOrCreatePage(yesterdayPageTitle, _settingsDailyPages.TemplateDailyPageTitle, 2);
        }

        public void GotoThisWeekPage ()
        {
            _templatePageCreator.GotoOrCreatePage(_settingsDailyPages.ThisWeekPageTitle(), _settingsDailyPages.TemplateWeeklyPageTitle);
        }

    }
}
