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
            var embedLinkToModelPage = ona.OneNoteLinkToPageIdWithExtra(modelPageId,extraId:modelPageId);
            SmartTagElementInDocument.Value = String.Format(hyperlinkFormatter, embedLinkToModelPage, "#") + SmartTagElementInDocument.Value.Substring(1);
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }
        private readonly static string hyperlinkFormatter = "<a href=\"{0}\">{1}</a>";

        public void SetLinkToPageId(OneNoteApp ona, string hierarchyElementId, string pageContentId = "")
        {
            SetLink(ona, new Uri(ona.GetHyperLinkToObject(hierarchyElementId, pageContentId)));
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

        public static string FullTextFromElementText(string elementText)
        {
            var augmenetedfullTextMatch = augmentedSmartTagMatcher.Match(elementText);
            var isAugmentedMatch = augmenetedfullTextMatch.Success;

            if (!isAugmentedMatch)
            {
                var fullTextMatch = Regex.Match(elementText,  notAugmentedSmartTagFullTextPattern);
                var fullText = fullTextMatch.Groups[1].Value;
                return fullText;
            }

            var augmentedfullText = "#" + augmenetedfullTextMatch.Groups["tagName"].Value + " "+ augmenetedfullTextMatch.Groups["fullText"];
            return augmentedfullText;
        }

        public static SmartTag FromElement(XElement element, OneNotePageCursor cursor, XDocument pageContent)
        {
            var elementText = element.Value;
            var extraIdMatch = augmentedSmartTagMatcher.Match(elementText);
            var fullText = FullTextFromElementText(elementText);

            return new SmartTag()
            {
                IsComplete = Regex.IsMatch(elementText, isCompleteMatcher),
                ModelPageId = extraIdMatch.Success ? extraIdMatch.Groups["extraId"].Value : "",
                FullText = fullText,
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
            var isRegularMatch = Regex.IsMatch(elementText, notAugmentedSmartTagPattern);
            var isAugmentedMatch = augmentedSmartTagMatcher.IsMatch(elementText);
            return isRegularMatch || isAugmentedMatch;
        }

        // the full text of the smart tag is harder to do without  a proper parser.
        // as a hack - we'll go from starting with a # to the end of elementText.
        private static string notAugmentedSmartTagPattern = "^(#.+) ";
        private static string notAugmentedSmartTagFullTextPattern = "^(#.+)";

        // Hack, assume the extraID as a pageId.
        //  I've seen pageID's of 86 and 87, so adding so putting buffer on both sides.

        private readonly static string _augmentedSmartTagPattern = "extraId=(?<extraId>.{85,90}}).*>#</a>.*>(?<tagName>.*)</a> (?<fullText>.+)";
        private readonly static Regex augmentedSmartTagMatcher = new Regex(_augmentedSmartTagPattern, RegexOptions.Singleline);


        private readonly static string isCompleteMatcher = "text-decoration:line-through"; 

        public void SetCompleted(OneNoteApp ona)
        {
            var currentText = this.SmartTagElementInDocument.Value;
            this.SmartTagElementInDocument.Value = String.Format("<span style='text-decoration:line-through'>{0}</span>", currentText);
            ona.OneNoteApplication.UpdatePageContent(PageContent.ToString());
        }
        public void AddEntryToModelPage(OneNoteApp ona, string text)
        {
            // TBD: It's possiblet the model page got deleted, handle that gracefully.
            var modelPageContent = ona.GetPageContentAsXDocument(ModelPageId);

            // NOTE: This is a big perf hit - figure out how to refactor.
            var page = OneNoteApp.XMLDeserialize<Page>(modelPageContent.ToString());

            DumbTodo.AddToPage(ona,modelPageContent, text, DateTime.Now);
        }
    }
}