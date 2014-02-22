using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OnenoteCapabilities;
using OneNoteObjectModel;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class PageTests
    {
        public OneNoteApp ona;

        // GRR - I don't understant why I can't open hierarchy, clearly I'm missing something.
        /*
        private static readonly string CurrentAssemblyPath = Path.GetDirectoryName(new Uri(Assembly.GetAssembly(typeof (OneNoteApp)).CodeBase).LocalPath);
        private static readonly string TestNoteBookPath = Path.Combine(CurrentAssemblyPath, @"..\..\..\TestNoteBooks\");
        */
        private TemporaryNoteBookHelper _templateNotebook;
        private TemporaryNoteBookHelper _dailyPagesNotebook;
        private OnenoteCapabilities.Settings settings;
        private DailyPages dailyPages;


        [TestFixtureSetUp]
        public void Setup()
        {
            ona = new OneNoteApp();
            _templateNotebook = new TemporaryNoteBookHelper(ona);
            _dailyPagesNotebook = new TemporaryNoteBookHelper(ona);

            settings = new Settings()
            {
                    TemplateNotebook = _templateNotebook.Get().name,
                    DailyPagesNotebook =  _dailyPagesNotebook.Get().name
            };

            // create template structure.
            var templateSection  = ona.CreateSection(_templateNotebook.Get(), settings.TemplateSection);
            ona.CreatePage(templateSection, settings.TemplateDailyPageTitle);
            ona.CreatePage(templateSection, settings.TemplateWeeklyPageTitle);

            // create DailyPages Section
            var dailySection = ona.CreateSection(_dailyPagesNotebook.Get(), settings.DailyPagesSection);
            ona.CreatePage(dailySection, "Parent Week");

            // Instantiate dailyPages
            dailyPages = new DailyPages(settings);
        }

        [Test]
        public void CreateWeek()
        {
            dailyPages.GotoThisWeekPage();
            // verify week page is created.
            var pagesNotebook = ona.GetNotebooks().Notebook.First(n => n.name == _dailyPagesNotebook.Get().name);
            var weekPage = ona.GetSections(pagesNotebook,true).First().Page.First(n => n.name == settings.ThisWeekPageTitle());
            Assert.That(weekPage.pageLevel, Is.EqualTo(1.ToString()));
        }
        [Test]
        public void CreateDay()
        {
            dailyPages.GotoTodayPage();
            // verify week page is created.
            var pagesNotebook = ona.GetNotebooks().Notebook.First(n => n.name == _dailyPagesNotebook.Get().name);
            var todayPage = ona.GetSections(pagesNotebook,true).First().Page.First(n => n.name == settings.TodayPageTitle());
            Assert.That(todayPage.pageLevel, Is.EqualTo(2.ToString()));
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            _templateNotebook.Dispose();
            _dailyPagesNotebook.Dispose();
        }
    }
}
