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
        public DailyPages(Settings settings)
        {
            this.settings = settings;
        }

        public Settings settings;

        public OneNoteApp ona = new OneNoteApp();
            
        // Goto page for today, create from template if it doesn't exist yet. 
        public void GotoTodayPage()
        {
            var pageTemplateForDay = ona.GetNotebooks().Notebook.First(n=>n.name == settings.TemplateNotebook )
                                .PopulatedSections(ona).First(s=>s.name == settings.TemplateSection)
                                .Page.First(p=>p.name == settings.TemplateDailyPageTitle);
            
            var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
                                .PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
            
            if (sectionForDailyPages.Page.Any(p=>p.name == settings.TodayPageTitle()))
            {
                Console.WriteLine("Today's template ({0}) has already been created,going to it",settings.TodayPageTitle);
            }
            else
            {
                var todaysPage = ona.ClonePage(sectionForDailyPages,pageTemplateForDay,settings.TodayPageTitle());
                Console.WriteLine("Created today's template page ({0}).",settings.TodayPageTitle);
                
                // Indent page because it will be folded into a weekly template.
                todaysPage.pageLevel = "2";
                ona.UpdatePage(todaysPage);
                
                // reload section since we modified the tree. 
                sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
                                .PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);    
            }

            var today = sectionForDailyPages.Page.First(p=>p.name == settings.TodayPageTitle());
            ona.OneNoteApplication.NavigateTo(today.ID);	

        }

        // TODO: Merge with Daily Page
        public void GotoThisWeekPage()
        {
            var pageTemplateForWeek = ona.GetNotebooks().Notebook.First(n=>n.name == settings.TemplateNotebook )
                                .PopulatedSections(ona).First(s=>s.name == settings.TemplateSection)
                                .Page.First(p=>p.name == settings.TemplateWeeklyPageTitle);
            
            var sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
                                .PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);     
            
            if (sectionForDailyPages.Page.Any(p=>p.name == settings.ThisWeekPageTitle()))
            {
                Console.WriteLine("This Week's template ({0}) has already been created,going to it",settings.TodayPageTitle);
            }
            else
            {
                var thisWeeksPage = ona.ClonePage(sectionForDailyPages,pageTemplateForWeek,settings.ThisWeekPageTitle());
                Console.WriteLine("Creating this week's template page ({0}).",settings.TodayPageTitle);
                ona.UpdatePage(thisWeeksPage);
                
                // reload section since we modified the tree. 
                sectionForDailyPages = ona.GetNotebooks().Notebook.First(n=>n.name == settings.DailyPagesNotebook)
                                .PopulatedSections(ona).First(s=>s.name == settings.DailyPagesSection);    
            }

            var thisWeek = sectionForDailyPages.Page.First(p=>p.name == settings.ThisWeekPageTitle());
            ona.OneNoteApplication.NavigateTo(thisWeek.ID);
        }
    }
}
