using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using OpenXMLTools.Interfaces;
using System.Linq;
using System.Text;

namespace OpenXMLTools.Word.OpenXmlElements
{
    public class ParagraphDecorator : Paragraph, IOpenXmlElementExtended
    {
        #region Fields

        private static int[] _numberedListIds = new int[] { 2, 3 };
        private static int[] _bulletListIds = new int[] { 1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };

        protected Paragraph _paragraph;

        #endregion

        #region Constructors

        public ParagraphDecorator(OpenXmlElement paragraph)
        {
            _paragraph = paragraph as Paragraph;
        }

        #endregion

        #region Public methods

        public Paragraph GetParagraph()
        {
            return _paragraph;
        }

        public string GetElementName(int nameLength)
        {
            StringBuilder result = new StringBuilder();
            StringBuilder text = new StringBuilder(_paragraph.InnerText);
            //foreach (var element in _paragraph.ChildElements)
            //{

            //    var runDescendants = element.Descendants<Run>();
            //    foreach (var runDescendant in runDescendants)
            //    {
            //        //hyperlink hidden text ignore
            //        if (WordTools.HasWebHiddenRunProperties(runDescendant))
            //            break;

            //        var textDescendants = runDescendant.Descendants<Text>();
            //        foreach (var textDescendant in textDescendants)
            //        {
            //            text.Append(textDescendant.Text);
            //       } 
            //    }
            //}
            result = WordTools.GetWordsFromTextElements(text, nameLength);

            return result.ToString();
        }

        public ElementType GetElementType()
        {
            if (IsNumberingList())
                return ElementType.NumberedList;
            else if (IsBulletList())
                return ElementType.BulletList;
            else if (IsHyperlink())
                return ElementType.Hyperlink;
            else
                return ElementType.Paragraph;
        }

        #endregion

        #region Private methods

        private bool IsNumberingList()
        {
            bool result = false;
            result = _numberedListIds.Any(b => b == WordTools.GetNumberingId(_paragraph));

            return result;
        }

        private bool IsBulletList()
        {
            bool result = false;
            result = _bulletListIds.Any(b => b == WordTools.GetNumberingId(_paragraph));

            return result;
        }

        private bool IsHyperlink()
        {
            bool result = false;

            var hyperlinkDescendants = _paragraph.Descendants<Hyperlink>();
            result = hyperlinkDescendants?.Count() > 0;

            return result;
        }

        #endregion
    }
}
