using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    // Put in for Minimum Viable Product Evaluation -full refactor requierd.
    public class PeopleAgendaSmartTagProcessor : ISmartTagProcessor
    {
        private SettingsPeoplePages settings;
        private Section peopleSection;
        private OneNoteApp ona;

        public PeopleAgendaSmartTagProcessor(OneNoteApp ona, SettingsPeoplePages settings)
        {
            this.ona = ona;
            this.settings = settings;
            this.peopleSection = ona.GetNotebook(settings.PeoplePagesNotebook)
                .PopulatedSection(ona, settings.PeoplePagesSection);
        }

        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            // todo test if I'm in the section for a person.
            if (peopleSection.ID != cursor.SectionId)
            {
                return false;
            }

            return (st.TagName() == "agenda");
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            var currentPage = ona.GetPageContent(cursor.PageId);
            // 2) Find my parent page
            var personName = currentPage.name.Split(':')[0];
            var personDate = currentPage.name.Split(':')[1];

            if (personDate != DateTime.Today.ToShortDateString())
            {
                smartTagAugmenter.AddContentAfterSmartTag(smartTag, pageContent ,String.Format("Augment only supported for a today's meeting page. {0} is not valid for #agenda",currentPage.name));
                return;
            }

            // HACK: refresh section  - need a better answer for this. 
            this.peopleSection = ona.GetNotebook(settings.PeoplePagesNotebook).PopulatedSection(ona, settings.PeoplePagesSection);

            // TOTAL Hack - find my parent page
            var personNextPage = peopleSection.Page.First(p => p.name == personName + ":Next");

            var personPageNextContent = ona.GetPageContentAsXDocument(personNextPage);
            var personPageCurrentContent = pageContent;

            // 3) Get data copies from my parent.
            var nextMyAction = personPageNextContent.DescendantNodes() .OfType<XElement>() .Where(e => e.Name.LocalName == "Table") .First();
            var nextThierAction = personPageNextContent.DescendantNodes() .OfType<XElement>() .Where(e => e.Name.LocalName == "Table") .Skip(1) .First();


            // Remove completed items from the agenda list. 

            var currentMyAction = personPageCurrentContent.DescendantNodes() .OfType<XElement>() .Where(e => e.Name.LocalName == "Table") .First();
            var currentThierAction = personPageCurrentContent.DescendantNodes() .OfType<XElement>() .Where(e => e.Name.LocalName == "Table") .Skip(1) .First();

            // todo strip completed items when importing the agenda. 

            RowsFromTable(currentMyAction).First().AddAfterSelf(RowsFromTable(nextMyAction).Where(RowContainsIncompleteTasks));
            RowsFromTable(currentThierAction).First().AddAfterSelf(RowsFromTable(nextThierAction).Where(RowContainsIncompleteTasks));

            ona.OneNoteApplication.UpdatePageContent(pageContent.ToString());
        }

        private bool RowContainsIncompleteTasks(XElement arg)
        {
            var tag = arg.DescendantNodes().OfType<XElement>().FirstOrDefault(e => e.Name.LocalName == "Tag");
            if (tag == default(XElement))
            {
                // no tag present in row. 
                return false;
            }

            if (tag.Attribute("completed").Value == "false")
            {
                return true;
            }
            return false;
        }

        private IEnumerable<XElement> RowsFromTable(XElement table)
        {
            return table.DescendantNodes().OfType<XElement>().Where(e => e.Name.LocalName == "Row");
        }
    }
}