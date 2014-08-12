using System.Linq;
using System.Xml.Linq;

namespace OnenoteCapabilities
{
    /// <summary>
    ///  Display in-line help for smarttags.
    /// </summary>
    public class HelpSmartTagProcessor:ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return st.TagName() == "help";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            var helpLines = smartTagAugmenter.GetAllHelpLines().ToList();

            // reverse the helpLines as we add elements after each other, so the first element shows last.
            helpLines.Reverse();

            foreach (var helpLine in helpLines)
            {
                smartTag.AddContentAfter(helpLine);
            }

        }

        public string HelpLine()
        {
            return "<b>#help</b> returns help";
        }
    }
}
