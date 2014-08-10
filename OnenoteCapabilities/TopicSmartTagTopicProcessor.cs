using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class TopicSmartTagTopicProcessor: ISmartTagProcessor
    {
        private OneNoteApp ona;
        private TemplatePageCreator templatePageCreater;
        private SettingsTopicPages settings;

        // TODO move from settingDailyPages to SettingTopicPages
        public TopicSmartTagTopicProcessor(OneNoteApp ona, SettingsTopicPages settings)
        {
            this.ona = ona;
            this.templatePageCreater = new TemplatePageCreator(ona, settings.TemplateNotebook, settings.TemplateSection, settings.TopicNotebook, settings.TopicSection);
            this.settings = settings;
        }
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            // Topic Processor should be added last because it will always make a smarttag.
            return true;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            // Create Topic Page of name 
            var topicPageName = smartTag.TagName();
            var topicPage = templatePageCreater.CreatePageIfNotExists(topicPageName,settings.TopicTemplateName,1);
            var topicPageContent = ona.GetPageContentAsXDocument(topicPage);

            DumbTodo.AddToPageFromDateEnableSmartTag(ona, topicPageContent,smartTag);

            if (OneNoteApp.IsSamePage (pageContent, topicPageContent))
            {
                // HACK: When the smartTag is on the current dailyPage, we put the TODO on the peoplePageContent, but then we put the link on pageContent, which doesn't have the changes to PageContent.
                pageContent = topicPageContent;
            }

            smartTag.SetLink(ona,templatePageCreater.SectionForPages(), topicPageName);
        }

        public string HelpLine()
        {
            return "<b>#topic</b> add a SmartTodo for the topic from the rest of this line";
        }
    }
}