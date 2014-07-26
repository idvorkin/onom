using System.Linq;
using System.Xml.Linq;
using NUnit.Framework;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class SmartTodoTests
    {
        private TemporaryNoteBookHelper smartTagNoteBook;
        private OneNoteApp ona;
        private XDocument pageContent;

        [SetUp]
        public void Setup()
        {
            this.ona = new OneNoteApp();
            smartTagNoteBook = new TemporaryNoteBookHelper(ona, "SmartTag");
            this.pageContent = smartTagNoteBook.CreatePage(new SmartTodoTestPageContent(), "Default Page");
        }

        [Test]
        public void Enumerate()
        {
            var smartTodos = SmartTodo.Get(pageContent);
            Assert.That(smartTodos.Count(),Is.EqualTo(3));
            Assert.That(smartTodos.Count(st=>!st.IsProcessed),Is.EqualTo(3));
            Assert.That(smartTodos.Count(st=>st.IsComplete),Is.EqualTo(1));
        }

        [Test]
        public void MarkComplete()
        {
            var smartTodos = SmartTodo.Get(pageContent);
            smartTodos.ToList().ForEach(st=>st.SetProcessed(ona));
            smartTodos = SmartTodo.Get(pageContent);
            Assert.That(smartTodos.Count(),Is.EqualTo(3));
            Assert.That(smartTodos.Count(st=>st.IsProcessed),Is.EqualTo(3));
        }

        [TearDown]
        public void TearDown()
        {
            smartTagNoteBook.Dispose();
        }
    }
}