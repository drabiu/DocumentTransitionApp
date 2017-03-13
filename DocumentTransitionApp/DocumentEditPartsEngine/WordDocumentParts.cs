using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools;
using System;
using System.Collections.Generic;
using System.IO;

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

        public static bool IsSupportedType(OpenXmlElement element)
        {
            bool isSupported = false;
            isSupported = element is Paragraph;
            //|| element is Wordproc.Picture
            //|| element is Wordproc.Drawing
            //|| element is Wordproc.Table;

            return isSupported;
        }
    }

    public class WordDocumentParts : IDocumentParts
    {
        int _paragraphCounter = 0;
        public List<PartsSelectionTreeElement> Get(Stream file, Predicate<OpenXmlElement> supportedType)
        {
            List<PartsSelectionTreeElement> documentElements = new List<PartsSelectionTreeElement>();
            byte[] byteArray = StreamTools.ReadFully(file);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(mem, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    for (int index = 0; index < body.ChildElements.Count; index++)
                    {
                        var element = body.ChildElements[index];
                        documentElements.AddRange(CreatePartsSelectionTreeElements(element, index, supportedType));
                    }
                }
            }

            return documentElements;
        }

        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            return Get(file, el => WordDocumentPartAttributes.IsSupportedType(el));
        }

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, int id, Predicate<OpenXmlElement> supportedType)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (supportedType(element))
            {
                PartsSelectionTreeElement elementToAdd;
                if (element is Paragraph)
                {
                    string elementId = (element as Paragraph).ParagraphId ?? WordDocumentPartAttributes.GetParagraphNoIdFormatter(_paragraphCounter);
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), 0, Helpers.ElementType.Paragraph);
                    _paragraphCounter++;
                }
                else
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), 0);

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
    }
}
