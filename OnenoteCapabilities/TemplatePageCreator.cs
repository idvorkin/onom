using System;
using System.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class TemplatePageCreator
    {
        private const int NO_INDENT_VALUE = -1;
        private readonly OneNoteApp ona;
        private readonly string templateNotebook;
        private readonly string templateSection;
        private readonly string pagesNotebook;
        private readonly string pagesSection;

        public TemplatePageCreator(OneNoteApp ona, string templateNotebook, string templateSection, string pagesNotebook, string pagesSection)
        {
            this.ona = ona;
            this.templateNotebook = templateNotebook;
            this.templateSection = templateSection;
            this.pagesNotebook = pagesNotebook;
            this.pagesSection = pagesSection;
        }

        public void GotoOrCreatePage (string pageTitle, string templateName, int indentValue)
        {
            var templateForDay = ona.GetNotebooks().Notebook.First(n => n.name == templateNotebook)
                .PopulatedSections(ona).First(s => s.name == templateSection)
                .Page.First(p => p.name == templateName);

            var sectionForPages = SectionForPages();

            bool isAnyPagesInSection = sectionForPages.Page != null;
            bool isPageExists =  isAnyPagesInSection && sectionForPages.Page.Any(p => p.name == pageTitle);
            if (!isPageExists)
            {
                // page does not exist.
                var newPage = ona.ClonePage(sectionForPages, templateForDay, pageTitle);

                // Indent page because it will be folded into a weekly template.
                if (indentValue != NO_INDENT_VALUE)
                {
                    newPage.pageLevel = indentValue.ToString();
                }
                else
                {
                    newPage.pageLevel = 1.ToString();
                }
                ona.UpdatePage(newPage);

                // reload section since we modified the tree. 
                sectionForPages = SectionForPages();
            }

            var reloadedNewPage = sectionForPages.Page.First(p => p.name == pageTitle);
            ona.OneNoteApplication.NavigateTo(reloadedNewPage.ID);
        }

        private Section SectionForPages()
        {
            var sectionForPages = ona.GetNotebooks().Notebook.First(n => n.name == pagesNotebook)
                .PopulatedSections(ona).First(s => s.name == pagesSection);
            return sectionForPages;
        }

        public void GotoOrCreatePage(string pageTitle, string templateName)
        {
            GotoOrCreatePage(pageTitle, templateName, NO_INDENT_VALUE);
        }
    }
}