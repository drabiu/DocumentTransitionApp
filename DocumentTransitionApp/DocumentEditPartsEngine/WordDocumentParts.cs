using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools;
using OpenXMLTools.Word.OpenXmlElements;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
            isSupported = element is Paragraph
                || element is Picture
                || element is Table
                || element is Drawing
                || element is Run;

            return isSupported;
        }
    }

    public class WordDocumentParts : IDocumentParts
    {
        int _paragraphCounter = 0;
        int _index = 0;

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

                    HashSet<OpenXmlElement> siblingsList = new HashSet<OpenXmlElement>();
                    foreach (var element in body.ChildElements)
                    {
                        //group list elements to one
                        if (siblingsList.Contains(element))
                            continue;

                        if (element is Paragraph && WordTools.IsListParagraph(element as Paragraph))
                        {
                            siblingsList = WordTools.GetAllSiblingListElements(element as Paragraph, body.ChildElements.ToList(), WordTools.GetNumberingId(element as Paragraph));
                        }

                        documentElements.AddRange(CreatePartsSelectionTreeElements(element, _index, supportedType, 0));
                        _index++;
                    }
                }
            }

            return documentElements;
        }

        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            return Get(file, el => WordDocumentPartAttributes.IsSupportedType(el));
        }

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, int id, Predicate<OpenXmlElement> supportedType, int indent)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (supportedType(element))
            {
                PartsSelectionTreeElement elementToAdd = null;
                if (element is Paragraph)
                {
                    ParagraphDecorator paragraphDecorator = new ParagraphDecorator(element);
                    string elementId = paragraphDecorator.GetParagraph().ParagraphId ?? WordDocumentPartAttributes.GetParagraphNoIdFormatter(_paragraphCounter);

                    if (paragraphDecorator.IsBulletList())
                        elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, paragraphDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.BulletList);
                    else if (paragraphDecorator.IsNumberingList())
                        elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, paragraphDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.NumberedList);
                    else
                        elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, paragraphDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.Paragraph);

                    _paragraphCounter++;
                }
                else if (element is Drawing)
                {
                    DrawingDecorator drawingDecorator = new DrawingDecorator(element);
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), drawingDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.Picture);
                }
                else if (element is Picture)
                {

                }
                else if (element is Table)
                {
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), (element as Table).LocalName, indent, Helpers.ElementType.Table);
                }

                if (elementToAdd != null)
                    result.Add(elementToAdd);

                if (element.HasChildren)
                {
                    foreach (var elmentChild in element.ChildElements)
                    {
                        result.AddRange(CreatePartsSelectionTreeElements(elmentChild, _index, supportedType, ++indent));
                    }

                }
            }

            return result;
        }

        private void GroupList()
        {

        }
    }
}
