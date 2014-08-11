using NUnit.Framework;
using OnenoteCapabilities;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class SmartTagXmlProcessingTests
    {
        public SmartTagXmlProcessingTests()
        {
            _cursor = new OneNotePageCursor();
        }

        private OneNotePageCursor _cursor;

        [Test] 
        public void TestNotTag()
        {
            var unprocessedTag = "notATag Tag2";
            Assert.That(!SmartTag.IsSmartTag(unprocessedTag));
        }
        [Test] 
        public void TestUnAugmenetedTag()
        {
            var tagText = "#hello world";
            Assert.That(SmartTag.IsSmartTag(tagText));
            Assert.AreEqual(tagText,SmartTag.FullTextFromElementText(tagText));
        }

        [Test]
        public void parseAugmentedTag()
        {
            var tagContent =
                "<a href=\"onenote:SmartTagStorage.one#Model%205844974a-21ac-4de5-83cc-9ec9aa362b0d&amp;section-id={D6DF7A74-EF20-476B-AC90-8D6D6BBBE4DC}&amp;page-id={2C9EA676-6657-4435-B4A1-E5AC6720AE34}&amp;end&amp;extraId={3EF0B7D7-1857-0633-04EE-212D80860CA2}{1}{E19543199418872411574120135711402651528027051}&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">#</a><a" +
                " href=\"onenote:Topics.one#hello&amp;section-id={25BAB527-01EA-417E-B55F-D52F8E5FB2EC}&amp;page-id={68DDB3B7-F9E5-455F-90B5-68E2823C5188}&amp;end&amp;base-path=https://d.docs.live.net/922579950926bf9e/Documents/BlogContentAndResearch\">hello</a> world";

            Assert.That(SmartTag.IsSmartTag(tagContent));
            Assert.AreEqual("#hello world",SmartTag.FullTextFromElementText(tagContent));

        }
    }
}