using System;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class PeoplePages
    {
        private SettingsPeoplePages _settings;
        public PeoplePages(SettingsPeoplePages settings)
        {
            this._settings = settings;
            _templatePageCreator = new TemplatePageCreator(
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
            var nextOrLastMeetingPage = _templatePageCreator.GetLastPageOfHeirarchyOrDefault(_settings.PersonNextTitle(person));
            if (nextOrLastMeetingPage == default(Page))
            {
                // Person next page doesn't exist - create it.
                GotoPersonNextPage(person);
                nextOrLastMeetingPage = _templatePageCreator.GetLastPageOfHeirarchyOrDefault(_settings.PersonNextTitle(person));
            }
            _templatePageCreator.GotoPage(nextOrLastMeetingPage.name);
        }

        public void GotoPersonCurrentMeetingPage (string person)
        {
            // TBD: Handle initial person page does not exist.

            // TBD: Page is inserted at bottom - now move to correct location.
            var nextOrLastMeetingPage = _templatePageCreator.GetLastPageOfHeirarchyOrDefault(_settings.PersonNextTitle(person));
            if (nextOrLastMeetingPage == default(Page))
            {
                // Person next page doesn't exist - create it.
                GotoPersonNextPage(person);
                nextOrLastMeetingPage = _templatePageCreator.GetLastPageOfHeirarchyOrDefault(_settings.PersonNextTitle(person));
            }

            _templatePageCreator.GotoOrCreatePageAfter(_settings.PersonMeetingTitle(person, System.DateTime.Now), _settings.TemplatePeopleMeetingTitle, 2, nextOrLastMeetingPage.name);
        }

    }
}
