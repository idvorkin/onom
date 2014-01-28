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
            var ona = new OneNoteObjectModel.OneNoteApp();
            var blankSections =
                ona.GetNotebooks()
                    .Notebook.SelectMany(n => n.PopulatedSections(ona))
                    .Where(s => s.IsDefaultUnmodified(ona));
            blankSections.ToList().ForEach(s => ona.OneNoteApplication.DeleteHierarchy(s.ID));
        }
    }
}

