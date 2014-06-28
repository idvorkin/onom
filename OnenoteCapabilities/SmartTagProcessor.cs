using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public interface ISmartTagProcessor
    {
        // TBD:Ensure matching functions don't overlap (easy source of bugs)
        bool ShouldProcess(SmartTag st);
        void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter);
    }

    public class PeopleSmartTagProcessor : ISmartTagProcessor
    {
        private SettingsPeoplePages settings;
        private Section peopleSection;

        public PeopleSmartTagProcessor(SettingsPeoplePages settings)
        {
            this.settings = settings;
        }

        public bool ShouldProcess(SmartTag st)
        {
            return settings.People().Contains(st.TagName());
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // TBD move to c'tor - but doesnt' yet get an smartTagAugmetor.

            this.peopleSection = smartTagAugmenter.ona.GetNotebooks()
                .Notebook.First(n => n.name == settings.PeoplePagesNotebook)
                .PopulatedSections(smartTagAugmenter.ona).First(s => s.name == settings.PeoplePagesSection);

            smartTagAugmenter.AddLinkToSmartTag(pageContent,smartTag,peopleSection,settings.PersonNextTitle(smartTag.TagName()));
        }
        
    }

    public class TwitterSmartTagProcessor: ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st)
        {
            return st.TagName().ToLower() == "tweet";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {


            // do nothing.
        }
    }
}
