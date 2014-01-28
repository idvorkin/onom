using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNoteObjectModel;

// TODO Add Test Cases.

namespace OnenoteCapabilities
{
    public class DailyPages
    {
        public DailyPages(Settings settings)
        {
            this.settings = settings;
        }

        public Settings settings;

        public OneNoteApp ona = new OneNoteApp();
        private const int NO_INDENT_VALUE = -1;

        public void GotoTodayPage()
        {
            string todayPageTitle = settings.TodayPageTitle();
            GotoOrCreateDayPage(todayPageTitle, settings.TemplateDailyPageTitle, 2);
        }

        public void GotoYesterday()
        {
            string yesterdayPageTitle = settings.DayPageTitleFromDate(DateTime.Now - TimeSpan.FromDays(1.0));
            GotoOrCreateDayPage(yesterdayPageTitle, settings.TemplateDailyPageTitle, 2);
        }

        public void GotoThisWeekPage ()
        {
            GotoOrCreateDayPage(settings.ThisWeekPageTitle(), settings.TemplateWeeklyPageTitle);
        }

        private void GotoOrCreateDayPage(string dayPageTitle, string templateName)
        {
            GotoOrCreateDayPage(dayPageTitle, templateName, NO_INDENT_VALUE);
        }

            // Goto page for today, create from template if it doesn't exist yet. 
        protected void GotoOrCreateDayPage (string dayPageTitle, string templateName, int indentValue)
        {
            var pageTemplateForDay = ona.GetNotebooks().Notebook.First(n => n.name == settings.TemplateNotebook)
                .PopulatedSections(ona).First(s => s.name == settings.TemplateSection)
                .Page.First(p => p.name == templateName);

            var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n => n.name == settings.DailyPagesNotebook)
                .PopulatedSections(ona).First(s => s.name == settings.DailyPagesSection);

            if (sectionForDailyPages.Page.Any(p => p.name == dayPageTitle))
            {
                Console.WriteLine(" page ({0}) has already been created,going to it", dayPageTitle);
            }
            else
            {
                var todaysPage = ona.ClonePage(sectionForDailyPages, pageTemplateForDay, dayPageTitle);
                Console.WriteLine("Created Page ({0} from template {1}).", dayPageTitle, templateName);

                // Indent page because it will be folded into a weekly template.
                if (indentValue != NO_INDENT_VALUE)
                {
                    todaysPage.pageLevel = indentValue.ToString();
                }
                ona.UpdatePage(todaysPage);

                // reload section since we modified the tree. 
                sectionForDailyPages = ona.GetNotebooks().Notebook.First(n => n.name == settings.DailyPagesNotebook)
                    .PopulatedSections(ona).First(s => s.name == settings.DailyPagesSection);
            }

            var today = sectionForDailyPages.Page.First(p => p.name == dayPageTitle);
            ona.OneNoteApplication.NavigateTo(today.ID);

        }
    }
}
