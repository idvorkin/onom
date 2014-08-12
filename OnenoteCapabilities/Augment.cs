using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public class OneNotePageCursor
    {
        public string PageId;
        public string SectionId;
        public string SectionGroupId;
        public string NotebookId;
    }
    public interface IPageAugmenter
    {
        void AugmentPage(XDocument pageContentInXml, OneNotePageCursor cursor );
    }


    public class Augmenter
    {
        public Augmenter(IEnumerable<IPageAugmenter> pageAugmentors)
        {
            this.pageAugmentors = pageAugmentors;
        }

        private IEnumerable<IPageAugmenter> pageAugmentors;

        /// <summary>
        /// Augment the page.
        /// </summary>

        public void AugmentCurrentPage()
        {
            var currentWindow = OneNoteApplication.Instance.InteropApplication.Windows.CurrentWindow;
            var currentLocation = new OneNotePageCursor()
            {
                PageId = currentWindow.CurrentPageId,
                SectionId = currentWindow.CurrentSectionId,
                SectionGroupId = currentWindow.CurrentSectionGroupId,
                NotebookId = currentWindow.CurrentNotebookId,
            };

            var pageContentInXML = OneNoteApplication.Instance.GetPageContentAsXDocument(currentLocation.PageId);

            foreach (var pageAugmentor in pageAugmentors)
            {
                pageAugmentor.AugmentPage(pageContentInXML, currentLocation);
            }
        }
    }
}
