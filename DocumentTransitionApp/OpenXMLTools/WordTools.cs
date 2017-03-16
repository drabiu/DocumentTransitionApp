using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Text;
using P = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXMLTools
{
    public class WordTools
    {
        #region static Public methods

        public static string GetElementName(OpenXmlElement element, int nameLength)
        {
            StringBuilder result = new StringBuilder();
            if (element is Paragraph)
            {
                var paragraph = element as Paragraph;
                if (paragraph.ChildElements.Any(ch => ch is Run))
                {
                    StringBuilder text = new StringBuilder();
                    foreach (Run run in paragraph.ChildElements.OfType<Run>())
                    {
                        text.Append(run.InnerText);
                    }

                    result = AppendTextFromElements(text, nameLength);
                }
            }
            else if (element is Table)
            {
            }
            else if (element is Picture)
            {

            }
            else if (element is Drawing)
            {
                var drawing = element as Drawing;
                StringBuilder text = new StringBuilder();
                foreach (var picture in drawing.Inline.Graphic.GraphicData.Elements<P.Picture>())
                {
                    foreach (var picProp in picture.Elements<P.NonVisualPictureProperties>())
                        text.Append(picProp.NonVisualDrawingProperties.Name);
                }

                result = AppendTextFromElements(text, nameLength);
            }

            return result.ToString();
        }

        #endregion

        private static StringBuilder AppendTextFromElements(StringBuilder text, int nameLength)
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
    }
}
