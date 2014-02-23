﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnenoteCapabilities
{
    // TODO Split this into different settings files.
    public class SettingsDailyPages:SettingsBase
    {
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

    public class SettingsPeoplePages:SettingsBase
    {
        public string TemplatePeopleNextTitle = "Person:Next";
        public string TemplatePeopleMeetingTitle = "Person:Meeting";
        public string PeoplePagesNotebook = "BlogContentAndResearch";
        public string PeoplePagesSection = "People";
        public string People = "SeanSe;IgorD;AlaksS;AmmonL;ToriS;SriRama;MaSudame;Demo1;Demo2";

        public string PersonMeetingTitle(string personName, DateTime now)
        {
            return String.Format("{0}:{1}",personName,now.ToShortDateString());
        }

        public string PersonNextTitle(string personName )
        {
            return String.Format("{0}:Next",personName);
        }
    }
}
