using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OneNoteObjectModel;

namespace OnenoteCapabilities
{
    [DebuggerDisplay("{FullText}")]
    public class InkTag
    {
        public string FullText;
        public XElement ElementForParent;

        public static IEnumerable<InkTag> Get(XDocument pageContentInXml)
        {
            var inkTagElements = pageContentInXml.DescendantNodes().OfType<XElement>().Where(IsInkTag);
            var inkTags = new List<InkTag>();
            foreach (var inkTagElement in inkTagElements)
            {
                var inkTag = FromElement(inkTagElement);
                inkTags.Add(inkTag);
            }
            return inkTags;
        }

        public void ToText()
        {
            ElementForParent.Descendants().Remove();
            OneNoteApp.AddContentAfter(FullText,ElementForParent);
        }

        private static InkTag FromElement(XElement inkTagElement)
        {
            Debug.Assert(IsInkTag(inkTagElement));

            var recognizedTextAttribute = inkTagElement.Attributes().FirstOrDefault(x => x.Name.LocalName == "recognizedText");
            var recognizedWord = recognizedTextAttribute.Value;

            // the next recognized words are the peers.
            var siblingRecognizedWordElements = inkTagElement.ElementsAfterSelf().Where(elem => RecognizedText(elem) != null);
            var recognizedWords = new List<string>() {recognizedWord};
            recognizedWords.AddRange(siblingRecognizedWordElements.Select(RecognizedText));
            var fullText = string.Join(" ", recognizedWords);

            // sometimes fullText ends up with a # tag, replace that.
            fullText = fullText.Replace("# ", "#");
            return new InkTag()
            {
                FullText = fullText,
                ElementForParent = inkTagElement.Parent
            };

        }

        private static string RecognizedText(XElement element)
        {
            if (element.Name.LocalName != "InkWord")
            {
                return null;
            }
            var recognizedTextAttribute = element.Attributes().FirstOrDefault(x => x.Name.LocalName == "recognizedText");
            if (recognizedTextAttribute == null)
            {
                return null;
            }
            var recognizedWord = recognizedTextAttribute.Value;
            return recognizedWord;
        }

        private static bool IsInkTag(XElement element)
        {
            var recognizedText = RecognizedText(element);
            return recognizedText != null && recognizedText.StartsWith("#");
        }
    }
}