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

    public class PeoplePages
    {
        private SettingsPeoplePages _settings;
        public PeoplePages(OneNoteApp ona, SettingsPeoplePages settings)
        {
            this._settings = settings;
            _templatePageCreator = new TemplatePageCreator(ona,
                    templateNotebook:settings.TemplateNotebook, 
                    templateSection:settings.TemplateSection, 
                    pagesNotebook:settings.PeoplePagesNotebook, 
                    pagesSection:settings.PeoplePagesSection);
        }

        private readonly TemplatePageCreator _templatePageCreator;

        public void GotoPersonNextPage(string person)
        {
            _templatePageCreator.GotoOrCreatePage(_settings.PersonNextTitle(person), _settings.TemplatePeopleNextTitle ,1);
        }

        public void GotoPersonPreviousMeetingPage(string person)
        {
            // TBD: Find Last Meeting - If Not Exist Short Curretn
            throw new NotImplementedException();
            /*
            string yesterdayPageTitle = _settings.DayPageTitleFromDate(DateTime.Now - TimeSpan.FromDays(1.0));
            _templatePageCreator.GotoOrCreatePage(yesterdayPageTitle, _settings.TemplateDailyPageTitle, 2);
            _templatePageCreator.GotoOrCreatePage(_settings.PersonMeetingTitle(person, System.DateTime.Now), _settings.TemplatePeopleMeetingTitle);
            */
        }

        public void GotoPersonCurrentMeetingPage (string person)
        {
            // TBD: Find Place To Insert Page.
            _templatePageCreator.GotoOrCreatePage(_settings.PersonMeetingTitle(person, System.DateTime.Now), _settings.TemplatePeopleMeetingTitle);
        }

    }
}
