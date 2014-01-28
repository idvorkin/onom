using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnenoteCapabilities
{
    public class Settings
    {
        public string TemplateNotebook = "Templates";
        public string TemplateSection = "Default";
        public string TemplateDailyPageTitle = "Daily";
        public string TemplateWeeklyPageTitle = "Weekly";
        public string DailyPagesNotebook = "BlogContentAndResearch";
        public string DailyPagesSection = "Current";
        // TodayPageTitle needs to be a functor as it depends on the day. 
        public string TodayPageTitle()
        {
            return DayPageTitleFromDate(DateTime.Now);
        }

        public string DayPageTitleFromDate(DateTime day)
        {
            return day.Date.ToShortDateString();
        }
        public string ThisWeekPageTitle() 
        {
               return "Week " + (DateTime.Now.Date - TimeSpan.FromDays((int) DateTime.Now.DayOfWeek - 1)).ToShortDateString();
        }
    }
}

