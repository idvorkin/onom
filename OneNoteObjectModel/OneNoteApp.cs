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
    public class OneNoteApp
    {
        private Application _oneNoteApplication = new Microsoft.Office.Interop.OneNote.Application();

        public Application OneNoteApplication
        {
            get { return _oneNoteApplication; }
        }

        public Notebooks GetNotebooks()
        {
            var serializer = new XmlSerializer(typeof (Notebooks));
            return XMLDeserialize<Notebooks>(GetHierarchy("",HierarchyScope.hsNotebooks));
        }

        // Return the hierarchy as a string (instead of an out paramater)
        public string GetHierarchy(string root, HierarchyScope hsScope)
        {
            string output;
            _oneNoteApplication.GetHierarchy(root, hsScope, out output);
            return output;
        }

        // Simple syntax object to XML string, but very inefficient.
        private string XMLSerialize<T>(T input)
        {
            var tempXML = new StringBuilder();
            new XmlSerializer(typeof (T)).Serialize(new XmlTextWriter(new StringWriter(tempXML)), input);
            return tempXML.ToString();
        }

        // I
        // Simple syntax XML string to object, but very inefficient.
        private T XMLDeserialize<T>(string input)
        {
            return (T) new XmlSerializer(typeof (T)).Deserialize(new XmlTextReader(new StringReader(input)));
        }

        public void CreateNoteBook(string path, string name)
        {
            // create the new notebook.
            var notebook = new Notebook {path = path, name = name, nickname = name};

            // add new notebook to notebooklist.
            var notebookList = XDocument.Parse(GetHierarchy("", HierarchyScope.hsNotebooks));
            notebookList.Root.Add(XDocument.Parse(XMLSerialize(notebook)).Root);

            // tell onenote about it.
            _oneNoteApplication.UpdateHierarchy(notebookList.ToString());


        }
    }
}
