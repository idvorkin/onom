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
        private Settings settings;

        public DailyPages(OneNoteApp ona, Settings settings)
        {
            this.settings = settings;
            _templatePageCreator = new TemplatePageCreator(ona,
                    templateNotebook:settings.TemplateNotebook, 
                    templateSection:settings.TemplateSection, 
                    pagesNotebook:settings.DailyPagesNotebook, 
                    pagesSection:settings.DailyPagesSection);
        }

        private readonly TemplatePageCreator _templatePageCreator;

        public void GotoTodayPage()
        {
            string todayPageTitle = settings.TodayPageTitle();
            GotoOrCreatePage(todayPageTitle, settings.TemplateDailyPageTitle, 2);
        }

        public void GotoYesterday()
        {
            string yesterdayPageTitle = settings.DayPageTitleFromDate(DateTime.Now - TimeSpan.FromDays(1.0));
            GotoOrCreatePage(yesterdayPageTitle, settings.TemplateDailyPageTitle, 2);
        }

        public void GotoThisWeekPage ()
        {
            _templatePageCreator.GotoOrCreatePage(settings.ThisWeekPageTitle(), settings.TemplateWeeklyPageTitle);
        }

        // Goto page for today, create from template if it doesn't exist yet. 
        protected void GotoOrCreatePage (string pageTitle, string templateName, int indentValue)
        {
            _templatePageCreator.GotoOrCreatePage(pageTitle, templateName, indentValue);
        }
    }
}
