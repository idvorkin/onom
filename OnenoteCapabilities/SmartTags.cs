using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
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
        public IEnumerable<SmartTag> GetSmartTagsOn(XDocument pageContent)
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
            var smartTags = GetSmartTagsOn(pageContent);
            foreach (var smartTag in smartTags.Where(st=>!st.IsProcessed()))
            {
                AddToModel(smartTag, pageContent);
                ProcessSmartTag(smartTag, pageContent);
            }
        }

        private void AddToModel(SmartTag smartTag, XDocument pageContent)
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

        public void AddLinkToSmartTag(XDocument pageContentInXML, SmartTag smartTag, Section linkSection, string linkPageName)
        {
            var smartTagElement =
                pageContentInXML.DescendantNodes()
                    .OfType<XElement>()
                    .First(r => r.Name.LocalName == "T" && r.Value.Contains(smartTag.ID.ToString()));

            var embedLinkToPersonPage = EmbedLinkToPage(linkPageName, linkSection);
            var personLink = string.Format(hyperlinkFormatter, embedLinkToPersonPage, smartTag.TagName());

            // TOTAL HACK NEEDS UNIT TESTS -- See AddIdToSmartTag
            smartTagElement.Value = smartTagElement.Value.Replace("#</a>"+smartTag.TagName(),"#</a>"+personLink);
            ona.OneNoteApplication.UpdatePageContent(pageContentInXML.ToString());
        }

        private void ProcessSmartTag(SmartTag smartTag, XDocument pageContent)
        {
            foreach (var tagProcessor in smartTagProcessors)
            {
                if (tagProcessor.ShouldProcess(smartTag))
                {
                    tagProcessor.Process(smartTag,pageContent, this);
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
        private readonly string smartTagFindFormatter = "ObjectId={0}\">.</a>";
        private readonly string infoFormatter = "<![CDATA[{0}]]>";

    }

    public interface ISmartTagProcessor
    {
        // TBD:Ensure matching functions don't overlap (easy source of bugs)
        bool ShouldProcess(SmartTag st);
        void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter);
    }

    public class PeopleSmartTagProcessor : ISmartTagProcessor
    {
        private SettingsPeoplePages settings;
        private Section peopleSection;

        public PeopleSmartTagProcessor(SettingsPeoplePages settings)
        {
            this.settings = settings;
        }

        public bool ShouldProcess(SmartTag st)
        {
            return settings.People().Contains(st.TagName());
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {
            // TBD move to c'tor - but doesnt' yet get an smartTagAugmetor.

            this.peopleSection = smartTagAugmenter.ona.GetNotebooks()
                .Notebook.First(n => n.name == settings.PeoplePagesNotebook)
                .PopulatedSections(smartTagAugmenter.ona).First(s => s.name == settings.PeoplePagesSection);

            smartTagAugmenter.AddLinkToSmartTag(pageContent,smartTag,peopleSection,settings.PersonNextTitle(smartTag.TagName()));
        }
        
    }

    public class TwitterSmartTagProcessor: ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st)
        {
            return st.TagName().ToLower() == "tweet";
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter)
        {


            // do nothing.
        }
    }
}
