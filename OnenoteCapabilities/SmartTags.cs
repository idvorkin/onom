using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Xml;
using System.Xml.Linq;
using OneNoteObjectModel;
using Page = OneNoteObjectModel.Page;

namespace OnenoteCapabilities
{
    public class SmartTag
    {
        // TODO: Consider adding a page TODO that helps me find hte ID in the page.
        public bool IsComplete;
        public string ModelPageId; // null if not set.
        public string FullText;
        public XElement SmartTagElementInDocument;
        public OneNotePageCursor CursorLocation;

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
    }

    public class SmartTagAugmenter:IPageAugmenter
    {
        public SmartTagAugmenter(OneNoteApp ona, SettingsSmartTags settings,IEnumerable<ISmartTagProcessor> smartTagProcessors)
        {
            Debug.Assert(smartTagProcessors != null);

            this.smartTagProcessors = smartTagProcessors;
            this.settings = settings;
            this.ona = ona;

            var smartTagNotebook = ona.GetNotebook(settings.SmartTagStorageNotebook);
            this.smartTagStorageSection = smartTagNotebook.PopulatedSection(ona,settings.SmartTagStorageSection);

            var templateNoteBook = ona.GetNotebook(settings.TemplateNotebook);
            this.smartTagTemplatePage = templateNoteBook.PopulatedSection(ona, settings.TemplateSection).GetPage(ona, this.settings.SmartTagTemplateName);
        }
        public IEnumerable<SmartTag> GetSmartTags(XDocument pageContent, OneNotePageCursor cursor)
        {
            var possibleSmartTags = pageContent.DescendantNodes().OfType<XElement>()
                .Where(r => r.Name.LocalName == "T" && !string.IsNullOrWhiteSpace(r.Value))
                .ToList();

            return possibleSmartTags
                .Select(t => new {element = t, elementText = t.Value})
                .Where(x => IsSmartTag(x.elementText))
                .Select(x=>SmartTagFromElement(x.elementText,x.element, cursor));
       }

        public static SmartTag SmartTagFromElement(string elementText, XElement element, OneNotePageCursor cursor)
        {
            var isCompleteMatcher = "text-decoration:line-through";
            
            var extraIdMatch = Regex.Match(elementText, SmartTagAugmenter.extraIdMatch);
            var fullTextMatch = Regex.Match(elementText, fullTextOfSmartTagMatcher);

            return new SmartTag()
            {
                IsComplete = Regex.IsMatch(elementText, isCompleteMatcher),
                ModelPageId = extraIdMatch.Success ? extraIdMatch.Groups[1].Value : "",
                FullText = fullTextMatch.Groups[1].Value,
                SmartTagElementInDocument = element,
                CursorLocation = cursor,
            };

        }
        public static bool IsSmartTag(string elementText)
        {
           const string smartTagRegExPattern = "^(#.+) ";
           var isRegularMatch = Regex.IsMatch(elementText, smartTagRegExPattern);
           var isAugmentedMatch = Regex.IsMatch(elementText, extraIdMatch);
           return isRegularMatch || isAugmentedMatch;
        }

        public void AugmentPage(OneNoteApp ona, XDocument pageContent, OneNotePageCursor cursor)
        {
            var smartTags = GetSmartTags(pageContent, cursor);
            foreach (var smartTag in smartTags.Where(st=>!st.IsProcessed()))
            {
                AddToModel(smartTag, pageContent);
                ProcessSmartTag(smartTag, pageContent, cursor);
            }
        }

        public void AddToModel(SmartTag smartTag, XDocument pageContent)
        {
            // create a new page to represent the smart tag. 
            var newModelPageName = string.Format("Model: {0}", Guid.NewGuid());
            var modelPage = ona.ClonePage(smartTagStorageSection,smartTagTemplatePage , newModelPageName );

            // put a hyper-link to the page on the '#'
            AddIdToSmartTag(pageContent,new List<SmartTag>(){smartTag}, newModelPageName, modelPage.ID);
                
            // TODO Refactor to somewhere else.
            // record human readable text on the Model Page.
            var modelPageContent = ona.GetPageContentAsXDocument(modelPage);

            // NOTE: This is a big perf hit - figure out how to refactor.
            var page = OneNoteApp.XMLDeserialize<Page>(pageContent.ToString());

            // TODO: Make this hyper-link.
            var creationText = String.Format("Instantiated model from tag '{1}' on page '{0}' with tag text:{2}", page.name, smartTag.TagName(),smartTag.TextAfterTag());
            DumbTodo.AddToPage(ona,modelPageContent, creationText, DateTime.Now);
        }

        // TODO: Refactor to be cleaner.
        public void AddIdToSmartTag(XDocument pageContentInXML, IEnumerable<SmartTag> smartTags, string modelPageName, string modelPageId)
        {
            var allSmartTagElements =
                pageContentInXML.DescendantNodes()
                    .OfType<XElement>()
                    .Where(r => r.Name.LocalName == "T" && IsSmartTag(r.Value));

            var smartTagElementsThatHaveNotBeenProcessed = allSmartTagElements.Where(st => !st.Value.Contains("ObjectId"));

            var embedLinkToModelPage = OneNoteLinkToPage(modelPageName, smartTagStorageSection, extraId:modelPageId);
            foreach (var smartTagElement in smartTagElementsThatHaveNotBeenProcessed)
            {
                var smartTag = smartTags.FirstOrDefault(s => s.FullText == Regex.Match(smartTagElement.Value, fullTextOfSmartTagMatcher).Groups[1].Value);
                if (smartTag == null) continue;

                // Update the smartTag with a GUID when we add it to the page.
                smartTag.ModelPageId = modelPageId;

                // Make the # a link to the model.
                // Add a '.' on the end to find the object ID. 
                smartTagElement.Value =
                    string.Format(hyperlinkFormatter, embedLinkToModelPage, "#") + smartTagElement.Value.Substring(1);
            }
            ona.OneNoteApplication.UpdatePageContent(pageContentInXML.ToString());
        }

        public static string OneNoteLinkToPage(string pageName, Section section, string extraId="")
        {
            var embedLinkToModelPage = string.Format("onenote:#{0}&base-path={1}", pageName, section.path);
            if (!String.IsNullOrEmpty(extraId))
            {
                embedLinkToModelPage += "&extraId=" + extraId;
            }
            return embedLinkToModelPage;
        }
        public void AddLinkToSmartTag(SmartTag smartTag, XDocument pageContentInXml, Uri link)
        {
            var smartTagElement =
                pageContentInXml.DescendantNodes()
                    .OfType<XElement>()
                    .First(r => r.Name.LocalName == "T" && r.Value.Contains(smartTag.ModelPageId.ToString()));

            var linkAsHTML = string.Format(hyperlinkFormatter, link.ToString(), smartTag.TagName());

            // TOTAL HACK NEEDS UNIT TESTS -- See AddIdToSmartTag
            smartTagElement.Value = smartTagElement.Value.Replace("#</a>"+smartTag.TagName(),"#</a>"+linkAsHTML);
            ona.OneNoteApplication.UpdatePageContent(pageContentInXml.ToString());
        }

        public void AddLinkToSmartTag(SmartTag smartTag, XDocument pageContentInXML, Section linkSection, string linkPageName)
        {
            var oneNoteLinkToPage = OneNoteLinkToPage(linkPageName, linkSection );
            AddLinkToSmartTag(smartTag,pageContentInXML,new Uri(oneNoteLinkToPage));
        }

        public void AddContentAfterSmartTag(SmartTag smartTag, XDocument pageContentInXml, string content)
        {
            var smartTagElement =
                pageContentInXml.DescendantNodes()
                    .OfType<XElement>()
                    .First(r => r.Name.LocalName == "T" && r.Value.Contains(smartTag.ModelPageId.ToString()));
            if (smartTagElement == null)
            {
                throw new Exception(string.Format("SmartTag not found. SmartTag ID '{0}.", smartTag.ModelPageId));
            }

            // NOTE: THIS ASSUMES ONENOTE XML STRUCTURE.
            var parentElement = smartTagElement.Parent;
            if (parentElement == null)
            {
                throw new Exception("Unexpected OneNote XML encountered. Expected smart tag to be embedded in an OE element.");
            }

            // Create a new paragraph to hold the content and insert it immediately after the smart tag.
            var newOEElement = CreateXElementForOneNote("OE");
            parentElement.AddAfterSelf(newOEElement);
            var newTElement = CreateXElementForOneNote("T");
            newOEElement.Add(newTElement);
            var newCDataElement = new XCData(content);
            newTElement.Add(newCDataElement);

            ona.OneNoteApplication.UpdatePageContent(pageContentInXml.ToString());
        }

        private XElement CreateXElementForOneNote(string localName)
        {
            XNamespace one = "http://schemas.microsoft.com/office/onenote/2013/onenote";
            return new XElement(one + localName,
                new XAttribute(XNamespace.Xmlns + "one", one));
        }

        private void ProcessSmartTag(SmartTag smartTag, XDocument pageContent, OneNotePageCursor cursor)
        {
            foreach (var tagProcessor in smartTagProcessors)
            {
                if (tagProcessor.ShouldProcess(smartTag, cursor))
                {
                    tagProcessor.Process(smartTag,pageContent, this, cursor);
                    break;
                }
            }
        }

        public Section smartTagStorageSection;
        public IEnumerable<ISmartTagProcessor> smartTagProcessors;
        private SettingsSmartTags settings;
        public OneNoteApp ona;
        private Page smartTagTemplatePage;
        private readonly string hyperlinkFormatter = "<a href=\"{0}\">{1}</a>";

        // the full text of the smart tag is harder to do without  a proper parser.
        // as a hack - we'll go from starting with a # to the end of elementText.
        private readonly static string fullTextOfSmartTagMatcher = "^(#.+)";

        // Hack, assume the extraID is always PageId.Length == 87. 
        // This is brittle code, test it well.
        private readonly static string extraIdMatch = "#.*extraId=(.{87})\"";
    }
}
