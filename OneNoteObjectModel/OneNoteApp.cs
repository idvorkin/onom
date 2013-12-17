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
        private Application _oneNoteApplication = new Microsoft.Office.Interop.OneNote.Application();

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
        public string XMLSerialize<T>(T input)
        {
            var tempXML = new StringBuilder();
            new XmlSerializer(typeof (T)).Serialize(new XmlTextWriter(new StringWriter(tempXML)), input);
            return tempXML.ToString();
        }

        // I
        // Simple syntax XML string to object, but very inefficient.
        public T XMLDeserialize<T>(string input)
        {
            return (T) new XmlSerializer(typeof (T)).Deserialize(new XmlTextReader(new StringReader(input)));
        }

        public Notebook CreateNoteBook(string path, string name)
        {
            // create the new notebook.
            var notebook = new Notebook {path = path, name = name, nickname = name};

            // add new notebook to notebooklist.
            var notebookList = XDocument.Parse(GetHierarchy("", HierarchyScope.hsNotebooks));
            notebookList.Root.Add(XDocument.Parse(XMLSerialize(notebook)).Root);

            // tell onenote about it.
            _oneNoteApplication.UpdateHierarchy(notebookList.ToString());
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
            OneNoteApplication.UpdatePageContent(XMLSerialize(page));
            return GetPageContent(page.ID);
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
            string clonedPageContent;
            OneNoteApplication.GetPageContent(pageToClone.ID,out clonedPageContent,PageInfo.piAll);
            var pageToCloneAsXDoc = XDocument.Parse(clonedPageContent);

            // When cloning a page need to remove all object ID's as oneonte neeeds to write them out.
            pageToCloneAsXDoc.DescendantNodes() .OfType<XElement>() .ToList() .ForEach(x => x.Attributes().Where(a => a.Name == "objectID").Remove());

            // Remove title as we don't want to replace it.
            pageToCloneAsXDoc.DescendantNodes().OfType<XElement>().Where(x => x.Name == "{http://schemas.microsoft.com/office/onenote/2013/onenote}Title").Remove();

            // Create the new Page and write it to onenote.
            var newPage = CreatePage(section, title);
            var newPageXML = XMLDeserialize<Page>(pageToCloneAsXDoc.ToString());
            // update the XML as it still points to the page to clone.
            newPageXML.ID = newPage.ID;
            OneNoteApplication.UpdatePageContent(XMLSerialize(newPageXML));

            // Return the cloned page with content.
            return GetPageContent(newPage.ID);
        }
    }
    public static class ExtensionMethods
    {
        // These extension methods need a OneNoteApp so that they can be static, without making a singleton app (TODO consider a singleton pattern)
        public static IEnumerable<Section> PopulatedSections(this Notebook notebook, OneNoteApp onom)
        {
            return onom.GetSections(notebook);
        }
    }
}
