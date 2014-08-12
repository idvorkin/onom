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
            var app = new OneNoteApplication();
            app.GetNotebooks().Notebook.Select(notebook=>notebook.name).ToList().ForEach(Console.WriteLine);
        }
    }
}
