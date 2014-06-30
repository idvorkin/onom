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
using System.Xml;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class SmartTag
    {
        // TODO: Consider adding a page TODO that helps me find hte ID in the page.
        public bool IsComplete;
        public Guid ID; // null if not set.
        public string FullText;
        public XElement SmarTagElementInDocument;

        // When we've processed a smart-tag we set its GUID back to onenote.
        public bool IsProcessed()
        {
            return ID != Guid.Empty;
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
            this.smartTagProcessors = smartTagProcessors;
            this.settings = settings;
            this.ona = ona;
            this.smartTagStorageSection = ona.GetNotebooks() .Notebook.First(n => n.name == this.settings.SmartTagNotebook)
                .PopulatedSections(ona).First(s => s.name == this.settings.SmartTagStorageSection);
            this.smartTagTemplatePage = ona.GetNotebooks()
                    .Notebook.First(n => n.name == settings.TemplateNotebook)
                    .PopulatedSections(ona).First(s => s.name == settings.TemplateSection)
                    .Page.First(p => p.name == this.settings.SmartTagTemplateName);
        }
        public IEnumerable<SmartTag> GetUnProcessedSmartTags(XDocument pageContent)
        {
            // var pageContentInXML = XML
            var smartTagMatcher = new Regex(smartTagRegExPattern);
            var objectIdMatcher = "ObjectId=(.*)\"";
            var isCompleteMatcher = "text-decoration:line-through";
            
            var possibleSmartTags = pageContent.DescendantNodes().OfType<XElement>()
                .Where(r => r.Name.LocalName == "T" && !string.IsNullOrWhiteSpace(r.Value))
                .ToList();

            return possibleSmartTags
                .Select(t => new {element = t, elementText = t.Value, match = smartTagMatcher.Match(t.Value)})
                .Where(x => x.match.Success)
                .Select(x =>
                {
                    var objectIdMatch = Regex.Match(x.elementText, objectIdMatcher);
                    var fullTextMatch = Regex.Match(x.elementText, fullTextOfSmartTagMatcher);

                    return new SmartTag()
                    {
                        IsComplete = Regex.IsMatch(x.elementText, isCompleteMatcher),
                        ID = objectIdMatch.Success ? Guid.Parse(objectIdMatch.Groups[1].Value) : Guid.Empty,
                        FullText = fullTextMatch.Groups[1].Value,
                        SmarTagElementInDocument = x.element
                    };
                });
        }

        public void AugmentPage(OneNoteApp ona, XDocument pageContent)
        {
            var smartTags = GetUnProcessedSmartTags(pageContent);
            foreach (var smartTag in smartTags.Where(st=>!st.IsProcessed()))
            {
                AddToModel(smartTag, pageContent);
                ProcessSmartTag(smartTag, pageContent);
            }
        }

        public void AddToModel(SmartTag smartTag, XDocument pageContent)
        {

            // create a new page to represent the smart tag. 
            var modelPageName = smartTag.TagName() + "|" + DateTime.Now;
            ona.ClonePage(smartTagStorageSection,smartTagTemplatePage , modelPageName );
            // put a hyper-link to the page on the '#'
            AddIdToSmartTag(pageContent,new List<SmartTag>(){smartTag},modelPageName);
            // record info on that page.
        }

        public void AddIdToSmartTag(XDocument pageContentInXML, IEnumerable<SmartTag> smartTags, string modelPageName)
        {
            var allSmartTagElements =
                pageContentInXML.DescendantNodes()
                    .OfType<XElement>()
                    .Where(r => r.Name.LocalName == "T" && Regex.IsMatch(r.Value, smartTagRegExPattern));

            var smartTagElementsThatHaveNotBeenProcessed = allSmartTagElements.Where(st => !st.Value.Contains("ObjectId"));

            var embedLinkToModelPage = EmbedLinkToPage(modelPageName, smartTagStorageSection);
            foreach (var smartTagElement in smartTagElementsThatHaveNotBeenProcessed)
            {

                var smartTag = smartTags.FirstOrDefault(s => s.FullText == Regex.Match(smartTagElement.Value, fullTextOfSmartTagMatcher).Groups[1].Value);
                if (smartTag == null) continue;

                // Update the smartTag with a GUID when we add it to the page.
                smartTag.ID = Guid.NewGuid();

                // Make the # a link to the model.
                // Add a '.' on the end to find the object ID. 
                 smartTagElement.Value =
                    string.Format(hyperlinkFormatter, embedLinkToModelPage, "#") + smartTagElement.Value.Substring(1) + string.Format(embedSmartTagIdFormatter, smartTag.ID);
            }
            ona.OneNoteApplication.UpdatePageContent(pageContentInXML.ToString());
        }

        private string EmbedLinkToPage(string pageName, Section section)
        {
            var embedLinkToModelPage = string.Format("onenote:#{0}&base-path={1}", pageName, section.path);
            return embedLinkToModelPage;
        }
        public void AddLinkToSmartTag(SmartTag smartTag, XDocument pageContentInXml, Uri link)
        {
            var smartTagElement =
                pageContentInXml.DescendantNodes()
                    .OfType<XElement>()
                    .First(r => r.Name.LocalName == "T" && r.Value.Contains(smartTag.ID.ToString()));

            var personLink = string.Format(hyperlinkFormatter, link.ToString(), smartTag.TagName());

            // TOTAL HACK NEEDS UNIT TESTS -- See AddIdToSmartTag
            smartTagElement.Value = smartTagElement.Value.Replace("#</a>"+smartTag.TagName(),"#</a>"+personLink);
            ona.OneNoteApplication.UpdatePageContent(pageContentInXml.ToString());
        }

        public void AddLinkToSmartTag(SmartTag smartTag, XDocument pageContentInXML, Section linkSection, string linkPageName)
        {
            var embedLinkToPersonPage = EmbedLinkToPage(linkPageName, linkSection);
            AddLinkToSmartTag(smartTag,pageContentInXML,new Uri(embedLinkToPersonPage));
        }

        public void AddContentAfterSmartTag(SmartTag smartTag, XDocument pageContentInXml, string content)
        {
            var smartTagElement =
                pageContentInXml.DescendantNodes()
                    .OfType<XElement>()
                    .First(r => r.Name.LocalName == "T" && r.Value.Contains(smartTag.ID.ToString()));
            if (smartTagElement == null)
            {
                throw new Exception(string.Format("SmartTag not found. SmartTag ID '{0}.", smartTag.ID));
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

        private void ProcessSmartTag(SmartTag smartTag, XDocument pageContent)
        {
            foreach (var tagProcessor in smartTagProcessors)
            {
                if (tagProcessor.ShouldProcess(smartTag))
                {
                    tagProcessor.Process(smartTag,pageContent, this);
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
        private readonly string smartTagRegExPattern = "(#[a-zA-Z0-9]+) ";
        private readonly string fullTextOfSmartTagMatcher = "(#[a-zA-Z0-9\\s]+)";
        private readonly string embedSmartTagIdFormatter = "<a href=\"ObjectId={0}\">.</a>";
        private readonly string contentFormatter = "<one:OE><one:T><![CDATA[{0}]]></one:T></one:OE>";

    }
}
