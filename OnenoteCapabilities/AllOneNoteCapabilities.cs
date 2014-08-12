using System.Collections.Generic;
using System.Linq;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteMenu
{
    public class AllOneNoteCapabilities
    {
        public readonly SettingsPeoplePages SettingsPeoplePages = new SettingsPeoplePages();
        public readonly SettingsDailyPages SettingsDailyPages = new SettingsDailyPages();
        public readonly SettingsTopicPages SettingsTopicPages = new SettingsTopicPages();
        public readonly EraseEmpty EraseSections = new EraseEmpty();
        public readonly DailyPages DailyPages;
        public readonly PeoplePages PeoplePages;
        public readonly string[] ListOfPeople;
        public readonly Augmenter Augmenter;

        public AllOneNoteCapabilities ()
        {
            DailyPages = new DailyPages(SettingsDailyPages);
            PeoplePages = new PeoplePages(SettingsPeoplePages);
            ListOfPeople = SettingsPeoplePages.People().ToArray();
            // TBD: Look up a dependency injection mechanism.
            var smartTagProcessors = new List<ISmartTagProcessor>()
            {
                new HelpSmartTagProcessor(), 
                new TwitterSmartTagProcessor(), 
                new WikipediaSmartTagProcessor(), 
                new ConnectSmartTagProcessor(), 
                new PeopleSmartTagProcessor(SettingsPeoplePages), 
                new DailySmartTagProcessor(SettingsDailyPages),
                new PeopleAgendaSmartTagProcessor(SettingsPeoplePages),
                // Topic smarttag processor needs to go last as it will create a topic page for any un-processed tag.
                new TopicSmartTagTopicProcessor(SettingsTopicPages)
            };

            var smartTagAugmentor = new SmartTagAugmenter(new SettingsSmartTags(), smartTagProcessors);
            var pageAugmentors = new List<IPageAugmenter>
            {
                new InkTagAugmenter(), // inkTags must go first as they are processed by the smartTagAugmentor
                smartTagAugmentor, 
                new SmartTodoAugmenter()
            };
            Augmenter = new Augmenter(pageAugmentors);
        }
    }
}
