using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.OneNote;
using NUnit.Framework;
using NUnit.Framework.Constraints;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class ContentProcessorTests
    {
        public PropertyBag CreateOnePropertyBag(IEnumerable<string> strings)
        {
            var bag1 = new PropertyBag();
            bag1.Properties.Add("1", strings.ToList());
            return bag1;
        }

        [Test]
        public void TestCopyCtorIsDeep()
        {
            var bag1 = CreateOnePropertyBag("A;B".Split(';'));
            var bag2 = new PropertyBag(bag1);
            Assert.That(bag2.Properties.First().Value.Count == 2);
            bag1.Properties.First().Value.Add("C");
            Assert.That(bag2.Properties.First().Value.Count == 2);
            Assert.That(bag1.Properties.First().Value.Count == 3);
        }

        [Test]
        public void TestMerge()
        {
            var bag1 = CreateOnePropertyBag("A;B".Split(';'));
            var bag2 = CreateOnePropertyBag("C;D".Split(';'));
            var bag3 = bag1.Merge(new List<PropertyBag>() {bag2});
            Assert.That(bag3.Properties.First().Value.Count == 4);
        }
    }
}
