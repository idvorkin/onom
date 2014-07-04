using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Microsoft.Office.Interop.OneNote;
using OneNoteObjectModel;



namespace OneNoteObjectModel
{
    // A wrapper around the COM API's here: http://msdn.microsoft.com/en-us/library/office/gg649853(v=office.14).aspx
    public class OneNoteApp
    {
        private Application _oneNoteApplication = new Application();

        public Application OneNoteApplication
        {
            get { return _oneNoteApplication; }
        }

        public Notebooks GetNotebooks()
        {
            // XXX: Figure out how to deal with performance implictions of this.
            // Maybe make this a paramater.
            return XMLDeserialize<Notebooks>(GetHierarchy("", HierarchyScope.hsNotebooks));
        }

        public Notebook GetNotebook(string notebookName)
        {
            try
            {
                return GetNotebooks().Notebook.First(n => n.name == notebookName);
            }
            catch (InvalidOperationException e)
            {
                throw new InvalidOperationException(String.Format("Could not find Notebook:{0}",notebookName),e);
            }
        }

        public IEnumerable<Section> GetSections(Notebook notebook)
        {
            return GetSections(notebook, true);
        }

        // GetSections is slower if you enumerate pages. If Performance is a concern set includePages to false.

        public IEnumerable<Section> GetSections(Notebook notebook, bool includePages)
        {
            return XMLDeserialize<Notebook>(GetHierarchy(notebook.ID,   includePages? HierarchyScope.hsPages : HierarchyScope.hsSections)).Section;
        }



        // Return the hierarchy as a string (instead of an out paramater)
        public string GetHierarchy(string root, HierarchyScope hsScope)
        {
            string output;
            _oneNoteApplication.GetHierarchy(root, hsScope, out output);
            return output;
        }

        // Simple syntax object to XML string, but very inefficient.
        public static string XMLSerialize<T>(T input)
        {
            var tempXML = new StringBuilder();
            new XmlSerializer(typeof (T)).Serialize(new XmlTextWriter(new StringWriter(tempXML)), input);
            return tempXML.ToString();
        }

        // I
        // Simple syntax XML string to object, but very inefficient.
        public static T XMLDeserialize<T>(string input)
        {
            return (T) new XmlSerializer(typeof (T)).Deserialize(new XmlTextReader(new StringReader(input)));
        }

        public Notebook CreateNoteBook(string path, string name)
        {
            // create the new notebook.
            var notebook = new Notebook {path = path, name = name, nickname = name};

            // add new notebook to notebooklist.
            var notebookList = XDocument.Parse(GetHierarchy("", HierarchyScope.hsNotebooks));
            notebookList.Root.AddFirst(XDocument.Parse(XMLSerialize(notebook)).Root);

            // tell onenote about it.
            _oneNoteApplication.UpdateHierarchy(notebookList.ToString());
            // NOTE: Can't use GetNotebook as that searches by name, not by path. 
            return GetNotebooks().Notebook.First(n => n.path == path);
        }

        public Section CreateSection(Notebook notebook, string name)
        {
            var sectionList = XDocument.Parse(GetHierarchy(notebook.ID, HierarchyScope.hsSections));
            sectionList.Root.Add(XDocument.Parse(XMLSerialize(new Section {name = name})).Root);
            _oneNoteApplication.UpdateHierarchy(sectionList.ToString());
            return GetSections(notebook).First(s => s.name == name);
        }

        public IEnumerable<Page> GetPages(Section section)
        {
            return XMLDeserialize<Section>(GetHierarchy(section.ID, HierarchyScope.hsPages)).Page;
        }

        public Page CreatePage(Section section, string title)
        {
            string pageId = String.Empty;
            OneNoteApplication.CreateNewPage(section.ID, out pageId);
            var page = GetPageContent(pageId);
            (page.Title.OE.First().Items.First() as TextRange).Value = title;
            return UpdatePage(page);
        }
        public Page UpdatePage(Page page)
        {
            OneNoteApplication.UpdatePageContent(XMLSerialize(page));
            // return what the page is from onenote (with for example ID's filled in)
            return GetPageContent(page);
        }

        // Need a concept here something around page stubs vs page content. 
        public Page GetPageContent(Page page)
        {
            return GetPageContent(page.ID);
        }

        public XDocument GetPageContentAsXDocument(Page p)
        {
            return  GetPageContentAsXDocument (p.ID);
        }

        public XDocument GetPageContentAsXDocument(string pageId)
        {
            string pageContent;
            OneNoteApplication.GetPageContent(pageId,out pageContent);
            return XDocument.Parse(pageContent);
        }

        public Page GetPageContent(string PageId)
        {
            string pageContent;
            OneNoteApplication.GetPageContent(PageId,out pageContent);
            return XMLDeserialize<Page>(pageContent);
        }

        // Onenote syntax is really complex, I recommend cloning pages instead of creating them by hand.
        public Page ClonePage (Section section,Page pageToClone, string title)
        {
            // copy the source page.
            var pageToCloneAsXDoc = this.GetPageContentAsXDocument (pageToClone.ID);

            // When cloning a page need to remove all object ID's as oneonte neeeds to write them out.
            pageToCloneAsXDoc.DescendantNodes() .OfType<XElement>() .ToList() .ForEach(x => x.Attributes().Where(a => a.Name == "objectID").Remove());

            // Remove title as we don't want to replace it.
            pageToCloneAsXDoc.DescendantNodes().OfType<XElement>().Where(x => x.Name == "{http://schemas.microsoft.com/office/onenote/2013/onenote}Title").Remove();

            // Create the new Page and write it to onenote.
            var newPage = CreatePage(section, title);
            var newPageXML = XMLDeserialize<Page>(pageToCloneAsXDoc.ToString());

            // update the XML as it still points to the page to clone.
            newPageXML.ID = newPage.ID;

            // Update current time to be now
            newPageXML.dateTime = DateTime.Now;
            OneNoteApplication.UpdatePageContent(XMLSerialize(newPageXML));

            // Return the cloned page with content.
            return GetPageContent(newPage);
        }

        public static bool IsSamePage(XDocument lhs, XDocument rhs)
        {
            var lhsID = lhs.DescendantNodes().OfType<XElement>().Where(e => e.Name.LocalName == "Page").Attributes("ID").First();
            var rhsID = rhs.DescendantNodes().OfType<XElement>().Where(e => e.Name.LocalName == "Page").Attributes("ID").First();
            return lhsID.Value == rhsID.Value;
        }
    }
    public static class ExtensionMethods
    {
        // These extension methods need a OneNoteApp so that they can be static, without making a singleton app (TODO consider a singleton pattern)
        public static IEnumerable<Section> PopulatedSections(this Notebook notebook, OneNoteApp ona)
        {
            return ona.GetSections(notebook);
        }


        // PopulatedSections is significantly slower the PopulatedSection because it returns all pages.
        public static Section PopulatedSection(this Notebook notebook, OneNoteApp ona, string sectionName)
        {
            var sections =  ona.GetSections(notebook, true);
            try
            {
                return sections.First(s=>s.name == sectionName);
            }
            catch (InvalidOperationException e)
            {
                throw new InvalidOperationException(String.Format("Could not find Section:{0} in Notebook:{1}",sectionName, notebook.name),e);
            }
        }

        /// <summary>
        ///  Test if a section is the default section added by the onenote applications.
        /// <returns></returns>
        public static bool IsDefaultUnmodified(this Section s, OneNoteApp ona)
        {
            // f it doesn't start with default new section starting string it's not empty.
            const string newSectionStartingName = "New Section";
            if (!s.name.StartsWith(newSectionStartingName)) return false;

            // If has default name, and no pages, it's empty
            if (!s.Page.Any()) return true;

            // If has more then one page, it's not empty.  
            if (s.Page.Count() != 1) return false;

            // If there are no page items or title it's empty
            return s.Page.First().IsDefaultUnmodified(ona);
        }

        public static bool IsDefaultUnmodified(this Page p, OneNoteApp ona)
        {
            // Get page content just to be safe.
            var page = ona.GetPageContent(p);
            return page.Items == null && page.Title.OE.First().Items.OfType<TextRange>().All(x => x.Value == "");
        }
    }
}
