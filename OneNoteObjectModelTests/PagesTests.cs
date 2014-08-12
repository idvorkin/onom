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
    public class DailyPageTests
    {
        private TemporaryNoteBookHelper _templateNotebook;
        private TemporaryNoteBookHelper _dailyPagesNotebook;
        private OnenoteCapabilities.SettingsDailyPages _settingsDailyPages;
        private DailyPages dailyPages;


        [TestFixtureSetUp]
        public void Setup()
        {
            _templateNotebook = new TemporaryNoteBookHelper("DailyPagesTemplate");
            _dailyPagesNotebook = new TemporaryNoteBookHelper("DailyPages");

            _settingsDailyPages = new SettingsDailyPages()
            {
                    TemplateNotebook = _templateNotebook.Get().name,
                    DailyPagesNotebook =  _dailyPagesNotebook.Get().name
            };

            // create template structure.
            var templateSection  = OneNoteApplication.Instance.CreateSection(_templateNotebook.Get(), _settingsDailyPages.TemplateSection);
            OneNoteApplication.Instance.CreatePage(templateSection, _settingsDailyPages.TemplateDailyPageTitle);
            OneNoteApplication.Instance.CreatePage(templateSection, _settingsDailyPages.TemplateWeeklyPageTitle);

            // create DailyPages Section
            var dailySection = OneNoteApplication.Instance.CreateSection(_dailyPagesNotebook.Get(), _settingsDailyPages.DailyPagesSection);
            OneNoteApplication.Instance.CreatePage(dailySection, "Parent Week");

            // Instantiate dailyPages
            dailyPages = new DailyPages(_settingsDailyPages);
        }

        [Test]
        public void CreateWeek()
        {
            dailyPages.GotoThisWeekPage();
            // verify week page is created.
            var pagesNotebook = OneNoteApplication.Instance.GetNotebook(_dailyPagesNotebook.Get().name);
            var weekPage = pagesNotebook.PopulatedSection(_settingsDailyPages.DailyPagesSection).Page.First(n => n.name == _settingsDailyPages.ThisWeekPageTitle());
            Assert.That(weekPage.pageLevel, Is.EqualTo(1.ToString()));
        }
        [Test]
        public void CreateDay()
        {
            dailyPages.GotoTodayPage();
            // verify week page is created.
            var pagesNotebook = OneNoteApplication.Instance.GetNotebook(_dailyPagesNotebook.Get().name);
            var todayPage = pagesNotebook.PopulatedSection(_settingsDailyPages.DailyPagesSection).Page.First(n => n.name == _settingsDailyPages.TodayPageTitle());
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
