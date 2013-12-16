using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNoteObjectModel;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Serialization;



namespace ListNoteBooks
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Microsoft.Office.Interop.OneNote.Application();
            string heirXML;
            app.GetHierarchy("", HierarchyScope.hsNotebooks, out heirXML);
            var serializer = new XmlSerializer(typeof(Notebooks));
            var notebooks = (Notebooks)serializer.Deserialize(new System.IO.StringReader(heirXML));
            notebooks.Notebook.Select(notebook=>notebook.name).ToList().ForEach(Console.WriteLine);
        }
    }
}
