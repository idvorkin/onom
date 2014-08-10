using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OnenoteCapabilities
{
    public interface ISmartTagProcessor
    {
        bool ShouldProcess(SmartTag st, OneNotePageCursor cursor);
        ///
        void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor);

        /// <summary>
        /// The html help for the SmartTag.
        /// </summary>
        string HelpLine();
    }
}
