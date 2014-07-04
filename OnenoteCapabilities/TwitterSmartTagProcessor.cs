using System;
using System.Xml.Linq;
using Tweetinvi;

namespace OnenoteCapabilities
{
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
}