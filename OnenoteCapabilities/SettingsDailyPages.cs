using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteObjectModel;

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
        public string PeoplePagesSettingsPage = "PeopleSettings";
        private List<string> _CachedPeople = null;

        public SettingsPeoplePages()
        {
            
        }

        public IEnumerable<string> People()
        {
            if (_CachedPeople == null)
            {
                var settingsPage = OneNoteApplication.Instance.GetNotebook(PeoplePagesNotebook).PopulatedSection(PeoplePagesSection).GetPage(PeoplePagesSettingsPage);
                var content = OneNoteApplication.Instance.GetPageContentAsXDocument(settingsPage);

                var peopleNode = content.DescendantNodes().OfType<XElement>().First(e => e.Value.Trim() == "People:");
                // format is 
                // <T> People: </T>
                // <OE Children> 
                //      ...
                //     <T> Person </T> 
                // </OE Children> 
                _CachedPeople = (peopleNode.NextNode as XElement).DescendantNodes().OfType<XElement>().Where(e => e.Name.LocalName == "T").Select(e => e.Value.Trim()).ToList();
            }

            return _CachedPeople;
        }

        public string PersonMeetingTitle(string personName, DateTime now)
        {
            return String.Format("{0}:{1}",personName,now.ToShortDateString());
        }

        public string PersonNextTitle(string personName )
        {
            return String.Format("{0}:Next",personName);
        }
    }

    public class SettingsSmartTags : SettingsBase
    {
        public string SmartTagTemplateName = "SmartTag";
        public string SmartTagStorageNotebook = "BlogContentAndResearch";
        public string SmartTagModelSection = "SmartTagStorage";
    }

    public class SettingsTopicPages : SettingsBase
    {
        public string TopicNotebook = "BlogContentAndResearch";
        public string TopicTemplateName = "Topic";
        public string TopicSection = "Topics";
    }
}


