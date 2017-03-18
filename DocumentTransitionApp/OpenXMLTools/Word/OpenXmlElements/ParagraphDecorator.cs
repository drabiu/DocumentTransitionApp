using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Interfaces;
using System.Linq;
using System.Text;

namespace OpenXMLTools.Word.OpenXmlElements
{
    public class ParagraphDecorator : Paragraph, IOpenXmlElementExtended
    {
        private static int[] _numberedListIds = new int[] { 2, 3 };
        private static int[] _bulletListIds = new int[] { 1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };

        protected Paragraph _paragraph;

        public ParagraphDecorator(OpenXmlElement paragraph)
        {
            _paragraph = paragraph as Paragraph;
        }

        public Paragraph GetParagraph()
        {
            return _paragraph;
        }

        public string GetElementName(int nameLength)
        {
            StringBuilder result = new StringBuilder();
            if (_paragraph.ChildElements.Any(ch => ch is Run))
            {
                StringBuilder text = new StringBuilder();
                foreach (Run run in _paragraph.ChildElements.OfType<Run>())
                {
                    text.Append(run.InnerText);
                }

                result = WordTools.GetWordsFromTextElements(text, nameLength);
            }

            return result.ToString();
        }

        public bool IsNumberingList()
        {
            bool result = false;
            result = _numberedListIds.Any(b => b == WordTools.GetNumberingId(_paragraph));

            return result;
        }

        public bool IsBulletList()
        {
            bool result = false;
            result = _bulletListIds.Any(b => b == WordTools.GetNumberingId(_paragraph));

            return result;
        }
    }
}
