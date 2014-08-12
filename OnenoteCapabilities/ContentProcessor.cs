using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
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

        public static class PropertyBagExtensions
        {
            public static PropertyBag Merge(this PropertyBag first, IEnumerable<PropertyBag> bags)
            {
                var ret = new PropertyBag(first);
                bags.ToList().ForEach(bag =>
                {
                    foreach (var property in bag.Properties)
                    {
                        if (!ret.Properties.ContainsKey(property.Key))
                        {
                            ret.Properties.Add(property.Key,new List<string>());
                        }
                        ret.Properties[property.Key].AddRange(property.Value);
                    };
                });
                return ret;
            }

        }
        
        public class PropertyBag
        {
            // TODO :Should property bag include a title?
            // TODO: Value will grow into a list of more complex objects, but lets start with strings.

            // Store only the most complex representation, but can use the IsFunctions below to see if there are special case parsers/serializers available.
            public Dictionary<string, List<string>> Properties = new Dictionary<string, List<string>>();

            public PropertyBag(PropertyBag first)
            {
                foreach (var property in first.Properties)
                {
                    // TBD add a test to make sure this is a deep copy.
                    this.Properties.Add(property.Key, property.Value.ToList());
                }
            }

            public PropertyBag()
            {
                // TODO: Complete member initialization
            }

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
            public ContentProcessor(bool ignoreUnsupportedElementError=false)
            {
                this.IgnoreUnsupportedElementError = ignoreUnsupportedElementError;
            }

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
                    if (o is OneNoteObjectModel.Image)
                    {
                        var image = o as Image;
                        if (image.OCRData != null && image.OCRData.OCRText != null ) 
                        {
                            return new List<string>(){image.OCRData.OCRText};
                        }
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
                var oesAtTitleElement = oes.SkipWhile(i =>
                {
                    var oe = i.Items[0];
                    if (!(oe is TextRange)) return true;
                    var text = oe as TextRange;
                    return text.Value != title;
                });

                var isHasSecondElement = oesAtTitleElement.First().Items.Count() == 2;

                try
                {
                    // figured out via pinning watch in debugger.
                    var nestedTable = ((OneNoteObjectModel.OE) (oesAtTitleElement.ToArray()[0].OEChildren[0].Items[0])).Items[0] as Table;
                    if (nestedTable != null)
                    {
                        return nestedTable;
                    }
                }
                catch
                {
                    // nested table wasn't there.
                }
                var nextOEItem = oesAtTitleElement.Skip(1).First().Items[0];
                if (nextOEItem is Table)
                {
                    return nextOEItem as Table;
                }

                throw new InvalidDataException("Unable to find table in Content for title:"+title);
            }

            // Because it's very easy to work with flat lists/tables in onenote, a great way to do categorization is to have a single column table, with bold items being keys.
            // Thus a one column table encoding has the key as bold columns
            //  |<b> A </b> | a1|a2|a3 |<b> B </b>  | b1 | b2 b3 |

            // Will result in the following property bag

            // Key | {Values}
            //  A  | {a1, a2 , a3 }
            //  B  | {b1, b2 , b3 }



            public PropertyBag PropertyBagFromOneColumnPropertyTable(OneNoteObjectModel.Table table)
            {
                if (table.Columns.Count() != 1) 
                {
                    throw new ArgumentOutOfRangeException("Table is not 1 dimensional");
                }

                var rowTexts = table.Row.Select(r=>r.Cell[0].OEChildren.First().Items.First()).Cast<OE>().Select(oe=>(oe.Items.First() as TextRange).Value);

                // TODO: - having a single bold element makes this a category, need a better model. 

                var boldMatcher = new Regex("style='font-weight:bold'*>(.*)</span.*>");
                // TODO: Figure out how to handle this, maybe throw? 
                var currentKey = "FIRST-ROW-SHOULD-BE-BOLD-TO-MAKE-IT-THE-KEY";
                var propertyBag = new PropertyBag();
                foreach (var rowText in rowTexts)
                {
                    bool isKey =  boldMatcher.Match(rowText).Success;
                    if(isKey)
                    {
                        var keyValueFromRow = boldMatcher.Match(rowText).Groups[1].Value;
                        currentKey = keyValueFromRow;
                        continue;
                    }
                    if (!propertyBag.Properties.ContainsKey(currentKey))
                    {
                        propertyBag.Properties.Add(currentKey,new List<string>(){});
                    }
                    propertyBag.Properties[currentKey].Add(rowText);
                }
                return propertyBag;
            }


            public PropertyBag PropertyBagFromTwoColumnTable(OneNoteObjectModel.Table table)
            {
                if (table.Columns.Count() != 2)
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
                var content = OneNoteApplication.Instance.GetPageContent(page);
                // ASSUME: Outline is the outer box which contains child items.
                // ASSUME: OEChildren is always a list of OE
                var children = content.Items.OfType<Outline>().SelectMany(x => x.OEChildren);
                var oes = children.SelectMany(x => x.Items).Cast<OE>();
                var table = GetTableAfterTitle(tableTitle, oes);
                var propertyBag = PropertyBagFromTwoColumnTable(table);
                return propertyBag.Properties[rowTitle];
            }
    }
}
