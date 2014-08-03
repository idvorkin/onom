using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    [DebuggerDisplay("{FullText};Complete={IsComplete};Processed={IsProcessed()}")]
    public class SmartTag
    {
        public bool IsComplete;
        public string ModelPageId; // null if not set.
        public string FullText;
        public XElement SmartTagElementInDocument;
        public OneNotePageCursor CursorLocation;
        public XDocument PageContent { get; set; }

        // When we've processed a smart-tag we set its GUID back to onenote.
        public bool IsProcessed()
        {
            return ModelPageId != "";
        }
        public string TagName()
        {
            return this.FullText.Split(' ').First().Substring(1);
        }

        public string TextAfterTag()
        {
            return String.Join(" ", this.FullText.Split(' ').Skip(1));
        }

        public void SetLink(OneNoteApp ona, Uri uri)
        {
            var linkAsHTML = String.Format(hyperlinkFormatter, uri.ToString(), TagName());

            // TOTAL HACK NEEDS UNIT TESTS -- See SetId
            SmartTagElementInDocument.Value = SmartTagElementInDocument.Value.Replace("#</a>"+TagName(),"#</a>"+linkAsHTML);
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }
        public void SetId(OneNoteApp ona, string modelPageName, string modelPageId, Section smartTagModelSection)
        {
            // Update the smartTag with a GUID when we add it to the page.
            ModelPageId = modelPageId;

            // Make the # a link to the model.
            // Add a '.' on the end to find the object ID. 
            var embedLinkToModelPage = OneNoteApp.OneNoteLinkToPage(modelPageName, smartTagModelSection, extraId:modelPageId);
            SmartTagElementInDocument.Value = String.Format(hyperlinkFormatter, embedLinkToModelPage, "#") + SmartTagElementInDocument.Value.Substring(1);
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }
        private readonly string hyperlinkFormatter = "<a href=\"{0}\">{1}</a>";

        public void SetLink(OneNoteApp ona, Section linkSection, string linkPageName)
        {
            SetLink(ona, new Uri(OneNoteApp.OneNoteLinkToPage(linkPageName, linkSection)));
        }

        public static IEnumerable<SmartTag> Get(XDocument pageContent, OneNotePageCursor cursor)
        {
            var possibleSmartTags = pageContent.DescendantNodes().OfType<XElement>()
                .Where(r => r.Name.LocalName == "T" && !String.IsNullOrWhiteSpace(r.Value))
                .ToList();

            return possibleSmartTags
                .Where(e => IsSmartTag(e.Value))
                .Select<XElement, SmartTag>(e=>FromElement(e, cursor, pageContent));
        }

        public static SmartTag FromElement(XElement element, OneNotePageCursor cursor, XDocument pageContent)
        {
            var elementText = element.Value;
            
            var extraIdMatch = Regex.Match(elementText, (string) SmartTag.extraIdMatch);
            var fullTextMatch = Regex.Match(elementText, (string) fullTextOfSmartTagMatcher);

            return new SmartTag()
            {
                IsComplete = Regex.IsMatch(elementText, isCompleteMatcher),
                ModelPageId = extraIdMatch.Success ? extraIdMatch.Groups[1].Value : "",
                FullText = fullTextMatch.Groups[1].Value,
                SmartTagElementInDocument = element,
                CursorLocation = cursor,
                PageContent = pageContent,
            };
        }
        public void AddContentAfter(OneNoteApp ona, string content)
        {
            // NOTE: THIS ASSUMES ONENOTE XML STRUCTURE.
            var parentElement = SmartTagElementInDocument.Parent;
            if (parentElement == null)
            {
                throw new Exception("Unexpected OneNote XML encountered. Expected smart tag to be embedded in an OE element.");
            }

            OneNoteApp.AddContentAfter(content, parentElement);
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }

        public static bool IsSmartTag(string elementText)
        {
            const string smartTagRegExPattern = "^(#.+) ";
            var isRegularMatch = Regex.IsMatch(elementText, smartTagRegExPattern);
            var isAugmentedMatch = Regex.IsMatch(elementText, (string) extraIdMatch);
            return isRegularMatch || isAugmentedMatch;
        }

        // the full text of the smart tag is harder to do without  a proper parser.
        // as a hack - we'll go from starting with a # to the end of elementText.
        private readonly static string fullTextOfSmartTagMatcher = "^(#.+)";

        // Hack, assume the extraID as a pageId.
        //  I've seen pageID's of 86 and 87, so adding so putting buffer on both sides.

        // This is brittle code, test it well.
        private readonly static string extraIdMatch = "#.*extraId=(.{85,90})\"";
        private static string isCompleteMatcher = "text-decoration:line-through";

        public void SetCompleted(OneNoteApp ona)
        {
            var currentText = this.SmartTagElementInDocument.Value;
            this.SmartTagElementInDocument.Value = String.Format("<span style='text-decoration:line-through'>{0}</span>", currentText);
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }
    }
}