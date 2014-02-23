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
            var templatePage = ona.GetNotebooks().Notebook.First(n => n.name == templateNotebook)
                .PopulatedSections(ona).First(s => s.name == templateSection)
                .Page.First(p => p.name == templateName);

            var sectionForPages = SectionForPages();

            bool isAnyPagesInSection = sectionForPages.Page != null;
            bool isPageExists =  isAnyPagesInSection && sectionForPages.Page.Any(p => p.name == pageTitle);
            if (!isPageExists)
            {
                // page does not exist.
                var newPage = ona.ClonePage(sectionForPages, templatePage, pageTitle);

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

        public void GotoPage(string title)
        {

            var page = ona.GetNotebooks().Notebook.First(n => n.name == pagesNotebook)
                .PopulatedSections(ona).First(s => s.name == pagesSection)
                .Page.First(p => p.name == title);

            ona.OneNoteApplication.NavigateTo(page.ID);
        }

        private Section SectionForPages()
        {
            var sectionForPages = ona.GetNotebooks().Notebook.First(n => n.name == pagesNotebook)
                .PopulatedSections(ona).First(s => s.name == pagesSection);
            return sectionForPages;
        }
        // 
        // PageParent1  // PageParent1 returns PageParent1 since no children
        // PageParent2
        //    PageChild1
        //    PageChild2 // PageParent2 returns PageChild2 since it's the last of the children.
        //
        public Page GetLastPageOfHeirarchyOrDefault(string pageTitle)
        {
            var pages  = SectionForPages().Page.ToList();
            var parentPage = pages.FirstOrDefault(p => p.name == pageTitle);
            if (parentPage == default(Page))
            {
                return default(Page);
            }
            var possibleChildPages = pages.SkipWhile(p => p != parentPage).Skip(1);

            var childIndent = Int32.Parse(parentPage.pageLevel) + 1;

            var childPages = possibleChildPages.TakeWhile(p=>p.pageLevel == childIndent.ToString()).ToList() ;
            if (!childPages.Any())
            {
                // no children.
                return parentPage;
            }
            return childPages.Last();
        }

        public void GotoOrCreatePageAfter(string pageTitle, string templateName, int indentValue, string pageTitleToInsertAfter)
        {
            GotoOrCreatePage(pageTitle,templateName,indentValue);

            // Page created at the bottom of the notebook, now move it to the correct location.
            var sectionForPages = SectionForPages();

            var pagesList = sectionForPages.Page.ToList();

            var parentPage = pagesList.FirstOrDefault(p => p.name == pageTitleToInsertAfter);


            if (parentPage == default(Page)) 
            {
                return;
            }

            // take the page from the end and stick it after the parent
            var parentIndex = pagesList.IndexOf(parentPage);
            var newlyInsertedPage = pagesList.Last();
            pagesList.Insert(parentIndex+1,newlyInsertedPage);

            // set pages to everything but the last page.
            sectionForPages.Page = pagesList.Take(pagesList.Count - 1).ToArray();
            ona.OneNoteApplication.UpdateHierarchy(OneNoteApp.XMLSerialize(sectionForPages));
        }

        public void GotoOrCreatePage(string pageTitle, string templateName)
        {
            GotoOrCreatePage(pageTitle, templateName, NO_INDENT_VALUE);
        }
    }
}