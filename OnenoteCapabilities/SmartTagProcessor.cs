using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteObjectModel;
using Tweetinvi;

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

        public PeopleSmartTagProcessor(OneNoteApp ona, SettingsPeoplePages settings)
        {
            this.settings = settings;
            this.peopleSection = ona.GetNotebooks()
                .Notebook.First(n => n.name == settings.PeoplePagesNotebook)
                .PopulatedSections(ona).First(s => s.name == settings.PeoplePagesSection);
        }

        public bool ShouldProcess(SmartTag st)
        {
            return settings.People().Contains(st.TagName());
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            var personPageTitle = settings.PersonNextTitle(smartTag.TagName());
            smartTagAugmenter.AddLinkToSmartTag(smartTag, pageContent, peopleSection, personPageTitle );
        }
    }

    
    public class TopicSmartTagTopicProcessor: ISmartTagProcessor
    {
        private OneNoteApp ona;
        private TemplatePageCreator templatePageCreater;

        // TODO move from settingDailyPages to SettingTopicPages
        public TopicSmartTagTopicProcessor(OneNoteApp ona, SettingsDailyPages settings)
        {
            this.ona = ona;
            this.templatePageCreater = new TemplatePageCreator(ona, settings.TemplateNotebook, settings.TemplateSection, settings.DailyPagesNotebook, settings.DailyPagesSection);
        }
        public bool ShouldProcess(SmartTag st)
        {
            // Topic Processor should be added last because it will always make a smarttag.
            return true;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // Create Topic Page of name 
            var topicPageName = smartTag.TagName();
            var page = templatePageCreater.CreatePageIfNotExists(topicPageName,"Topic",1);
            var newPageContent = "";
            ona.OneNoteApplication.GetPageContent(page.ID,out newPageContent);
            AddDumbTodoToTopicPage(ona,XDocument.Parse(newPageContent),smartTag.TextAfterTag());
            smartTagAugmenter.AddLinkToSmartTag(smartTag,pageContent,templatePageCreater.SectionForPages(), topicPageName);
        }

        private void AddDumbTodoToTopicPage(OneNoteApp oneNoteApp, XDocument pageContentAsXml, string todo)
        {
            AddDumbTodoToTopicPage(oneNoteApp,pageContentAsXml,todo,DateTime.MinValue);
        }

        private void AddDumbTodoToTopicPage(OneNoteApp ona, XDocument pageContentAsXML, string todo, DateTime dueDate)
        {
            var rowTemplate = "<one:Row lastModifiedTime=\"2014-06-28T06:11:19.000Z\" xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"> " +
              "<one:Cell lastModifiedTime=\"2014-06-28T06:11:19.000Z\"  lastModifiedByInitials=\"ID\"> " +
                "<one:OEChildren> " +
                  "<one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:11:19.000Z\" lastModifiedTime=\"2014-06-28T06:11:19.000Z\" alignment=\"left\" quickStyleIndex=\"1\"> " +
                    "<one:Tag index=\"0\" completed=\"{2}\" disabled=\"false\" creationDate=\"2014-06-28T06:11:25.000Z\" completionDate=\"2014-06-28T06:27:01.000Z\" />"+
                    "<one:T><![CDATA[{0}]]></one:T> " +
                  "</one:OE> " +
                "</one:OEChildren> " +
              "</one:Cell> " +
              "<one:Cell lastModifiedTime=\"2014-06-28T06:11:13.000Z\" lastModifiedByInitials=\"ID\"> " +
                "<one:OEChildren> " +
                  "<one:OE authorResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" lastModifiedByResolutionID=\"&lt;resolutionId provider=&quot;Windows Live&quot; hash=&quot;XDgdaY/mTnjLm73zUG5SXQ==&quot;&gt;&lt;localId cid=&quot;922579950926bf9e&quot;/&gt;&lt;/resolutionId&gt;\" creationTime=\"2014-06-28T06:11:13.000Z\" lastModifiedTime=\"2014-06-28T06:11:13.000Z\" alignment=\"left\" quickStyleIndex=\"1\"> " +
                  "<one:T><![CDATA[{1}]]></one:T> " +
                  "</one:OE> " +
                "</one:OEChildren> " +
              "</one:Cell> " +
            "</one:Row>";

            bool completed=false;

            var row = string.Format(rowTemplate,todo,dueDate == DateTime.MinValue ? "": dueDate.ToShortDateString(),completed.ToString().ToLower());
            var rowAsXML = XDocument.Parse(row);

            // Add row after the first row (which is assumed to be a header in the first table)
            pageContentAsXML.DescendantNodes().OfType<XElement>().First(e => e.Name.LocalName=="Row").AddAfterSelf(rowAsXML.Root);
            ona.OneNoteApplication.UpdatePageContent(pageContentAsXML.ToString());
        }
    }

    /// <summary>
    ///  An example of using smartTags to push data out of OneNote. This tweets to the onenotehat twitter account.
    /// </summary>
    public class TwitterSmartTagProcessor: ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st)
        {
            return st.TagName().ToLower() == "tweet";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            TweetString(smartTag.TextAfterTag());
            smartTagAugmenter.AddLinkToSmartTag(smartTag,pageContent,new Uri("http://twitter.com/onenotehat"));
        }
        public static bool TweetString(string text)
        {
            // These credentials are hard-coded to the onenotehat account - to implement correctly 
            // get a user token and store it in the settings.

            TwitterCredentials.SetCredentials(
                consumerKey: "EwUvZ1yCZkmqyHAvx5eATGgmu",
                consumerSecret: "h4uJb2ySeCAJV4dc3QouzF3EfNkzJkVUJ0WR5NjB1J8ysZTMI7",
                userAccessToken: "2589375188-PcbQckRyzacxiKj2h1HgbVHpbMt1jbjSP6gCDM6",
                userAccessSecret: "GpPdq1GsYW9S1VPv82pqQvwHU3UPFWnYUKl9nqnrcMUH8"
                );

            var tweetText = String.Format("OneNoteLabs TweetSmarTag:{0}", text);
            var newTweet = Tweet.CreateTweet(tweetText);
            newTweet.Publish();
            return newTweet.IsTweetPublished;
        }
    }
}
