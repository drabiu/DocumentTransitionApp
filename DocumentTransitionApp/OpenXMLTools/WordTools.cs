using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                    result.Append("[Par]: ");
                    StringBuilder text = new StringBuilder();
                    foreach (Run run in paragraph.ChildElements.OfType<Run>())
                    {
                        text.Append(run.InnerText);
                    }

                    var listWords = text.ToString().Split(default(char[]), StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in listWords)
                    {
                        result.Append(string.Format("{0} ", word));
                        if (result.Length > nameLength)
                            break;
                    }

                    result.Remove(result.Length - 1, 1);
                }
            }
            else if (element is Table)
            {


            }
            else if (element is Picture)
            {

            }
            else if (element is Drawing)
            { }

            return result.ToString();
        }

        #endregion
    }
}
