using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentEditPartsEngine.Interfaces;
using System.IO;
using DocumentFormat.OpenXml;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentEditPartsEngine
{
    public class WordDocumentParts : IDocumentParts
    {
        private class WordDocumentPartAttributes
        {
            public const int MaxNameLength = 30;
        }

        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            List<PartsSelectionTreeElement> documentElements = new List<PartsSelectionTreeElement>();
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(file, true))
            {
                Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
                for (int index = 0; index < body.ChildElements.Count; index++)
                {
                    var element = body.ChildElements[index];
                    documentElements.AddRange(CreatePartsSelectionTreeElements(element, index));
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
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), (element as Wordproc.Paragraph).ParagraphId, GetElementName(element), 0);
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
                    result.Append("Paragraph: ");
                    foreach (Wordproc.Run run in paragraph.ChildElements.OfType<Wordproc.Run>())
                    {
                        result.Append(run.InnerText);
                        if (result.Length > WordDocumentPartAttributes.MaxNameLength)
                            break;
                    }
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
