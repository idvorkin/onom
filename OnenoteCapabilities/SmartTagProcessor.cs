using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
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

    public class DailySmartTagProcessor : ISmartTagProcessor
    {
        private SettingsDailyPages settings;
        private OneNoteApp ona;
        private Section dailySection;
        private DumbTodo _dumbTodo;

        public DailySmartTagProcessor(OneNoteApp ona, SettingsDailyPages settings)
        {
            this.ona = ona;
            this.settings = settings;
            this.dailySection = ona.GetNotebooks()
                .Notebook.First(n => n.name == settings.DailyPagesNotebook)
                .PopulatedSections(ona).First(s => s.name == settings.DailyPagesSection);
            _dumbTodo = new DumbTodo();

        }
        public bool ShouldProcess(SmartTag st)
        {
            return (st.TagName() == "today");
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            var todayPageTitle = settings.DayPageTitleFromDate(DateTime.UtcNow);
            var dailyPage = dailySection.Page.First(p => p.name == todayPageTitle);
            var dailyPageContent = ona.GetPageContentAsXDocument(dailyPage);

            // HACK: Need to find the table of interest with a better method.
            var hackTableToAddTasksTo = 4;

            _dumbTodo.AddDumbTodoToPage(smartTagAugmenter.ona, dailyPageContent,smartTag.TextAfterTag(),tableOnPage:hackTableToAddTasksTo);

            if (OneNoteApp.IsSamePage (pageContent, dailyPageContent))
            {
                // HACK: When the smartTag is on the current dailyPage, we put the TODO on the dailyPage, but then we put the link on pageContent, which doesn't have the changes to PageContent.
                pageContent = dailyPageContent;
            }

            smartTagAugmenter.AddLinkToSmartTag(smartTag, pageContent, dailySection, todayPageTitle);
        }
    }

    public class PeopleSmartTagProcessor : ISmartTagProcessor
    {
        private SettingsPeoplePages settings;
        private Section peopleSection;
        private DumbTodo _dumbTodo;
        private OneNoteApp ona;

        public PeopleSmartTagProcessor(OneNoteApp ona, SettingsPeoplePages settings)
        {
            this.ona = ona;
            this.settings = settings;
            _dumbTodo = new DumbTodo();
            this.peopleSection = ona.GetNotebooks()
                .Notebook.First(n => n.name == settings.PeoplePagesNotebook)
                .PopulatedSections(ona).First(s => s.name == settings.PeoplePagesSection);
        }

        public bool ShouldProcess(SmartTag st)
        {
            if (settings.People().Contains(personFromPersonTag(st))) return true;
            return false;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // TODO: Create Person Page if not Exists.
            var personPageTitle = settings.PersonNextTitle(personFromPersonTag(smartTag));

            // get PersonPage 

            var peoplePage = peopleSection.Page.First(p => p.name == personPageTitle);
            var peoplePageContent = ona.GetPageContentAsXDocument(peoplePage);

            const int toPersonTableCountOnPage = 0;
            const int fromPersonTableCountOnPage = 1;

            var tableOnPage = IsFromPerson(smartTag) ? fromPersonTableCountOnPage : toPersonTableCountOnPage;
            var parsedDate = HumanizedDateParser.ParseDateAtEndOfSentance(smartTag.TextAfterTag());
            if (parsedDate.Parsed)
            {
                _dumbTodo.AddDumbTodoToPage(smartTagAugmenter.ona, peoplePageContent, parsedDate.SentanceWithoutDate, dueDate:parsedDate.date, tableOnPage:tableOnPage);
            }
            else
            {

                _dumbTodo.AddDumbTodoToPage(smartTagAugmenter.ona, peoplePageContent, smartTag.TextAfterTag(), tableOnPage: tableOnPage);
            }

            if (OneNoteApp.IsSamePage (pageContent, peoplePageContent))
            {
                // HACK: When the smartTag is on the current dailyPage, we put the TODO on the peoplePageContent, but then we put the link on pageContent, which doesn't have the changes to PageContent.
                pageContent = peoplePageContent;
            }


            smartTagAugmenter.AddLinkToSmartTag(smartTag, pageContent, peopleSection, personPageTitle);
        }

        public bool IsFromPerson(SmartTag smartTag)
        {
            return smartTag.TagName().StartsWith("from");
        }

        public string personFromPersonTag (SmartTag smartTag)
        {
            if (IsFromPerson(smartTag))
            {
                return smartTag.TagName().Substring(4);
            }
            else
            {
                return smartTag.TagName();
            }
        }
    }

    
    public class TopicSmartTagTopicProcessor: ISmartTagProcessor
    {
        private OneNoteApp ona;
        private TemplatePageCreator templatePageCreater;
        private readonly DumbTodo _dumbTodo;

        // TODO move from settingDailyPages to SettingTopicPages
        public TopicSmartTagTopicProcessor(OneNoteApp ona, SettingsDailyPages settings)
        {
            this.ona = ona;
            this.templatePageCreater = new TemplatePageCreator(ona, settings.TemplateNotebook, settings.TemplateSection, settings.DailyPagesNotebook, settings.DailyPagesSection);
            _dumbTodo = new DumbTodo();
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
            var topicPage = templatePageCreater.CreatePageIfNotExists(topicPageName,"Topic",1);
            var topicPageContent = ona.GetPageContentAsXDocument(topicPage);

            var parsedDate = HumanizedDateParser.ParseDateAtEndOfSentance(smartTag.TextAfterTag());

            if (parsedDate.Parsed)
            {
                _dumbTodo.AddDumbTodoToPage(ona, topicPageContent, parsedDate.SentanceWithoutDate, parsedDate.date);
            }
            else
            {
                _dumbTodo.AddDumbTodoToPage(ona, topicPageContent, smartTag.TextAfterTag());
            }


            smartTagAugmenter.AddLinkToSmartTag(smartTag,pageContent,templatePageCreater.SectionForPages(), topicPageName);
        }
    }

    /// <summary>
    ///  An example of using smartTags to push data out of OneNote. This tweets to the onenotehat twitter account.
    /// </summary>
    public class TwitterSmartTagProcessor : ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st)
        {
            return st.TagName().ToLower() == "tweet";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            TweetString(smartTag.TextAfterTag());
            smartTagAugmenter.AddLinkToSmartTag(smartTag, pageContent, new Uri("http://twitter.com/onenotehat"));
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

            var tweetText = String.Format("OneNoteLabs Tweet:{0}", text);
            var newTweet = Tweet.CreateTweet(tweetText);
            newTweet.Publish();
            return newTweet.IsTweetPublished;
        }
    }

    /// <summary>
    ///  An example of using smartTags to push data in to OneNote. 
    ///  This grabs the first paragraph of an article from wikipedia.
    /// </summary>
    public class WikipediaSmartTagProcessor : ISmartTagProcessor
    {
        private static readonly string ExtractUrlFormatter = @"http://en.wikipedia.org/w/api.php?format=xml&action=query&prop=extracts&titles={0}&redirects=true";
        private static readonly string ArticleUrlFormatter = @"http://en.wikipedia.org/wiki/{0}";
        private static readonly string SearchUrlFormatter = @"Wikipedia information not found for topic. <br /><a href='http://en.wikipedia.org/wiki/Special:Search?search={0}'>Search Wikipedia for '{0}'.</a>";
        private static readonly string FirstParagraphPattern = @"<p>.+<\/p>";

        public bool ShouldProcess(SmartTag st)
        {
            return st.TagName().Equals("info", StringComparison.InvariantCultureIgnoreCase);
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            var search = smartTag.TextAfterTag();
            var info = GetWikipediaExtract(search);

            // Insert the content.
            smartTagAugmenter.AddContentAfterSmartTag(smartTag, pageContent, info);

            // Make the smart tag a link to the wikipedia article.
            smartTagAugmenter.AddLinkToSmartTag(smartTag, pageContent,
                new Uri(string.Format(ArticleUrlFormatter, WebUtility.UrlEncode(search))));
        }

        private string GetWikipediaExtract(string search)
        {
            // Attempt to get the topic from the web using the exact search string.
            var url = string.Format(ExtractUrlFormatter, WebUtility.UrlEncode(search));
            var webClient = new WebClient();
            var downloadResult = webClient.DownloadString(url);

            // Load this into an XML document for easy parsing.
            var wikiXml = new XmlDocument();
            wikiXml.LoadXml(downloadResult);

            // Parse out the first paragraph.
            var extract = wikiXml.SelectSingleNode("//extract");
            if (extract != null)
            {
                var match = Regex.Match(extract.InnerText, FirstParagraphPattern);
                if (match != null)
                {
                    // Return the result text from the article.
                    return match.Groups[0].Value;
                }
            }

            // Couldn't get the information, so give a standard message and a link to search.
            return string.Format(SearchUrlFormatter, WebUtility.UrlEncode(search), search);
        }

    }
}
