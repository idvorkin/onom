using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    public interface IPageAugmenter
    {
        void AugmentPage(OneNoteApp ona, XDocument pageContentInXml);
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
            var currentPageId = ona.OneNoteApplication.Windows.CurrentWindow.CurrentPageId;
            string pageContent;
            ona.OneNoteApplication.GetPageContent(currentPageId,out pageContent);
            var pageContentInXML = XDocument.Parse(pageContent);
            foreach (var pageAugmentor in pageAugmentors)
            {
                pageAugmentor.AugmentPage(ona,pageContentInXML);
            }
        }
    }
}
