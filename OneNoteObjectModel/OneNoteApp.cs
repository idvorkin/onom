using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Office.Interop.OneNote;
using OneNoteObjectModel;

namespace OneNoteObjectModel
{
    public class OneNoteApp
    {
        public Application onenote = new Microsoft.Office.Interop.OneNote.Application();

        public Application Onenote
        {
            get { return onenote; }
        }

        public Notebooks GetNotebooks()
        {
            string hierarchyXmlOut;
            onenote.GetHierarchy("", HierarchyScope.hsNotebooks, out hierarchyXmlOut);
            var serializer = new XmlSerializer(typeof (Notebooks));
            return (Notebooks) serializer.Deserialize(new System.IO.StringReader(hierarchyXmlOut));
        }
    }
}
