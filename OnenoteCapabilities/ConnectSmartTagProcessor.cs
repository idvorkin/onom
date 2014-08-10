using System;
using System.Xml.Linq;

namespace OnenoteCapabilities
{
    /// <summary>
    /// Convert connect tags to mail:// tags
    /// </summary>
    public class ConnectSmartTagProcessor : ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return st.TagName() == "connect";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            // TBD: FIgure out what to do if not an email address.
            var mailLink = String.Format("mailto:{0}", smartTag.TextAfterTag());
            smartTag.SetLink(smartTagAugmenter.ona, new Uri(mailLink));
        }

        public string HelpLine()
        {
            return "<b>#connect</b> make a connection to the person";
        }
    }
}