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
        void AugmentPage(OneNoteApp ona, XDocument pageContentInXml, OneNotePageCursor cursor );
    }


    public class Augmenter
    {
        public Augmenter(OneNoteApp ona, IEnumerable<IPageAugmenter> pageAugmentors)
        {
            this.pageAugmentors = pageAugmentors;
            this.ona = ona;
        }

        private IEnumerable<IPageAugmenter> pageAugmentors;
        private OneNoteApp ona;

        /// <summary>
        /// Augment the page.
        /// </summary>

        public void AugmentCurrentPage()
        {
            var currentWindow = ona.OneNoteApplication.Windows.CurrentWindow;
            var currentLocation = new OneNotePageCursor()
            {
                PageId = currentWindow.CurrentPageId,
                SectionId = currentWindow.CurrentSectionId,
                SectionGroupId = currentWindow.CurrentSectionGroupId,
                NotebookId = currentWindow.CurrentNotebookId,
            };

            var pageContentInXML = ona.GetPageContentAsXDocument(currentLocation.PageId);

            foreach (var pageAugmentor in pageAugmentors)
            {
                pageAugmentor.AugmentPage(ona,pageContentInXML, currentLocation);
            }
        }
    }
}
