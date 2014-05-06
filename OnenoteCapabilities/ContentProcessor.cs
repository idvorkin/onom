using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
        // TODO: Add lots of tests 
        // TODO: Define good words for this.

        // A Property Bag is essentially an AST for data stored in OneNote.
        // Semantic meaning can be built on top of ASTs
        // ASTs are parsed from various formats.
        // ASTs ban be converted to various formats.
        // While we store the most complex AST, if an AST has special properties, there are simpler parser/converters available.
        
        public class PropertyBag
        {
            // TODO :Should property bag include a title?
            // TODO: Value will grow into a list of more complex objects, but lets start with strings.

            // Store only the most complex representation, but can use the IsFunctions below to see if there are special case parsers/serializers available.
            public Dictionary<string, List<string>> Properties = new Dictionary<string, List<string>>();

            // When property bags have the following shapes they can have different syntactic representatics.
            public bool IsShapeSingleProperty()
            {
                return Properties.Count() == 1;
            }
            public bool IsShapeAllSingleValueShape()
            {
                return Properties.Values.All(x => x.Count == 1);
            }
        }

        public class ContentProcessor
        {
            public ContentProcessor(OneNoteApp ona, bool ignoreUnsupportedElementError=false)
            {
                this.ona = ona;
                this.IgnoreUnsupportedElementError = ignoreUnsupportedElementError;
            }

            public readonly OneNoteApp ona;
            public bool IgnoreUnsupportedElementError;

            public List<string> OneNoteContentToList(IEnumerable<Object> objects)
            {
                return objects.SelectMany(o =>
                {

                    if (o is OEChildren)
                    {
                        return OneNoteContentToList((o as OEChildren).Items);
                    }
                    if (o is OE)
                    {
                        return OneNoteContentToList((o as OE).Items);
                    }
                    if (o is TextRange)
                    {
                        return new List<string>() { (o as TextRange).Value };
                    }
                    if (o is OneNoteObjectModel.Table)
                    {
                        // "skipping Table Processing".Dump();
                        return new List<string>();
                    }
                    if (o is OneNoteObjectModel.InkWord)
                    {
                        return new List<string>() { (o as InkWord).recognizedText ?? "" };
                    }
                    else 
                    {
                        if (!IgnoreUnsupportedElementError)
                        {
                            throw new NotImplementedException("Unable To Parse an Element Of Type:"+o.GetType());
                        }
                        return new List<string>();
                    }
                }
                ).ToList();
            }

            public Table GetTableAfterTitle(string title, IEnumerable<OE> oes)
            {
                // TODO Add Test if title isn't there.
                return 
                // Iterate to the table Title.
                oes.SkipWhile(i =>
                {
                    var oe = i.Items[0];
                    if (!(oe is TextRange)) return true;
                    var text = oe as TextRange;
                    return text.Value != title;
                })
                // skip the title Table
                .Skip(1)
                // get the next element - which is the Table Itself.
                .First().Items[0] as OneNoteObjectModel.Table;
            }

            public PropertyBag PropertyBagFromTwoColumnTable(OneNoteObjectModel.Table table)
            {
                if (!table.Row.All(r => r.Cell.Count() == 2))
                {
                    throw new ArgumentOutOfRangeException("Table is not 2 dimensional");
                }

                // TBD  HANDLE MORE THEN One OEChildren.
                // var c = 
                var keys = table.Row.Select(r => OneNoteContentToList(r.Cell[0].OEChildren)).Select(k => k.First());
                var values = table.Row.Select(r => OneNoteContentToList(r.Cell[1].OEChildren));
                var propertyBag = new PropertyBag();
                propertyBag.Properties = keys.Zip(values, (k, v) => new { k, v }).ToDictionary(z => z.k, z => z.v);
                return propertyBag;
            }

            public IEnumerable<string> GetTableRowContent(OneNoteObjectModel.Page page, string tableTitle, string rowTitle)
            {
                var content = ona.GetPageContent(page);
                // ASSUME: Outline is the outer box which contains child items.
                // ASSUME: OEChildren is always a list of OE
                var children = content.Items.OfType<Outline>().SelectMany(x => x.OEChildren);
                var oes = children.SelectMany(x => x.Items).Cast<OE>();
                return PropertyBagFromTwoColumnTable(GetTableAfterTitle(tableTitle,oes)).Properties[rowTitle];
            }
    }
}
