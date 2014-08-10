using System;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OnenoteCapabilities
{
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

        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return st.TagName().Equals("info", StringComparison.InvariantCultureIgnoreCase);
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            var search = smartTag.TextAfterTag();
            var info = GetWikipediaExtract(search);

            // Insert the content.
            smartTag.AddContentAfter(smartTagAugmenter.ona, info);

            // Make the smart tag a link to the wikipedia article.
            smartTag.SetLink(smartTagAugmenter.ona, new Uri(string.Format(ArticleUrlFormatter, WebUtility.UrlEncode(search))));
        }

        public string HelpLine()
        {
            return "<b>#info</b> get information for the rest of this line";
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