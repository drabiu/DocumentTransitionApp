using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentEditPartsEngine
{
    public static class WordDocumentPartAttributes
    {
        public const int MaxNameLength = 35;
        public const string ParagraphHasNoId = "noid:";

        public static string GetParagraphNoIdFormatter(int id)
        {
            return string.Format("{0}{1}", ParagraphHasNoId, id);
        }
    }

    public class WordDocumentParts : IDocumentParts
    {
        int _paragraphCounter = 0;
        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            List<PartsSelectionTreeElement> documentElements = new List<PartsSelectionTreeElement>();
            byte[] byteArray = StreamTools.ReadFully(file);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(mem, true))
                {
                    Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
                    for (int index = 0; index < body.ChildElements.Count; index++)
                    {
                        var element = body.ChildElements[index];
                        documentElements.AddRange(CreatePartsSelectionTreeElements(element, index));
                    }
                }
            }

            return documentElements;
        }

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, int id)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (IsSupportedType(element))
            {
                PartsSelectionTreeElement elementToAdd;
                if (element is Wordproc.Paragraph)
                {
                    string elementId = (element as Wordproc.Paragraph).ParagraphId ?? WordDocumentPartAttributes.GetParagraphNoIdFormatter(_paragraphCounter);
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, GetElementName(element), 0);
                    _paragraphCounter++;
                }
                else
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), GetElementName(element), 0);

                result.Add(elementToAdd);
                if (element.HasChildren)
                {
                    CreateChildrenPartsSelectionTreeElements(element);
                }
            }

            return result;
        }

        private IEnumerable<PartsSelectionTreeElement> CreateChildrenPartsSelectionTreeElements(OpenXmlElement element)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            return result;
        }

        private string GetElementName(OpenXmlElement element)
        {
            StringBuilder result = new StringBuilder();
            if (element is Wordproc.Paragraph)
            {
                var paragraph = element as Wordproc.Paragraph;
                if (paragraph.ChildElements.Any(ch => ch is Wordproc.Run))
                {
                    result.Append("[Par]: ");
                    StringBuilder text = new StringBuilder();
                    foreach (Wordproc.Run run in paragraph.ChildElements.OfType<Wordproc.Run>())
                    {
                        text.Append(run.InnerText);                                                        
                    }

                    var listWords = text.ToString().Split(default(char[]), StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in listWords)
                    {
                        result.Append(string.Format("{0} ", word));
                        if (result.Length > WordDocumentPartAttributes.MaxNameLength)
                            break;
                    }

                    result.Remove(result.Length - 1, 1);
                }
            }
            else if (element is Wordproc.Table)
            {


            }
            else if (element is Wordproc.Picture)
            {

            }
            else if (element is Wordproc.Drawing)
            { }

            return result.ToString();
        }

        private bool IsSupportedType(OpenXmlElement element)
        {
            bool isSupported = false;
            isSupported = element is Wordproc.Paragraph;
            //|| element is Wordproc.Picture
            //|| element is Wordproc.Drawing
            //|| element is Wordproc.Table;

            return isSupported;
        }
    }
}
