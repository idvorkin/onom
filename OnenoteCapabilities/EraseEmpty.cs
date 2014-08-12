using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class EraseEmpty
    {
        public void DeleteEmptySections()
        {
            var blankSections =
                OneNoteApplication.Instance.GetNotebooks()
                    .Notebook.SelectMany(n => n.PopulatedSections())
                    .Where(s => s.IsDefaultUnmodified());
            blankSections.ToList().ForEach(s => OneNoteApplication.Instance.InteropApplication.DeleteHierarchy(s.ID));
        }
    }
}

