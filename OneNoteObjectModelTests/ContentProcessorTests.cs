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
        [Test]
        public void TestInitialization()
        {
            var ona = new OneNoteApp();
            var cp = new ContentProcessor(ona);
            // TODO: Add real tests.
        }
    }
}
