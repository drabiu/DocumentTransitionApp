using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXMLTools
{
    public class WordTools
    {
        #region static Public methods

        public static StringBuilder GetWordsFromTextElements(StringBuilder text, int nameLength)
        {
            StringBuilder result = new StringBuilder();
            var listWords = text.ToString().Split(default(char[]), StringSplitOptions.RemoveEmptyEntries);
            foreach (var word in listWords)
            {
                result.Append(string.Format("{0} ", word));
                if (result.Length > nameLength)
                    break;
            }

            if (result.Length > 0)
                result.Remove(result.Length - 1, 1);

            return result;
        }

        public static HashSet<OpenXmlElement> GetAllSiblingListElements(Paragraph paragraph, List<OpenXmlElement> elements, int numberingId)
        {
            IList<OpenXmlElement> result = new List<OpenXmlElement>();
            if (GetNumberingId(paragraph) == numberingId)
            {
                result.Add(paragraph);
                var index = elements.FindIndex(e => e is Paragraph && (e as Paragraph).ParagraphId == paragraph.ParagraphId);
                foreach (var element in elements.Skip(index + 1))
                {
                    if (element is Paragraph && GetNumberingId(element as Paragraph) == numberingId)
                        result.Add(element);
                    else
                        break;
                }
            }

            return new HashSet<OpenXmlElement>(result);
        }

        public static int GetNumberingId(Paragraph paragraph)
        {
            int result = 0;
            var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            if (numberingProperties != null)
            {
                result = numberingProperties.NumberingId.Val?.Value ?? 0;
            }

            return result;
        }

        public static bool IsListParagraph(Paragraph paragraph)
        {

            var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            bool result = numberingProperties != null;

            return result;
        }

        public static bool HasWebHiddenRunProperties(Run run)
        {
            bool result = false;
            var runProperties = run.Descendants<RunProperties>();
            foreach (var runProp in runProperties)
            {
                var webHidden = runProp.ChildElements.OfType<WebHidden>();
                if (webHidden != null && webHidden.Count() > 0)
                    result = true;
            }

            return result;
        }

        #endregion
    }
}
