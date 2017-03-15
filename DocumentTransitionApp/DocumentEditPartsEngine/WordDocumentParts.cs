using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocumentEditPartsEngine
{
    public static class WordDocumentPartAttributes
    {
        public static int[] NumberedListIds = new int[] { 2, 3 };
        public static int[] BulletListIds = new int[] { 1 };

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
                || element is Drawing;

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

                    foreach (var element in body.ChildElements)
                    {
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
                PartsSelectionTreeElement elementToAdd;
                if (element is Paragraph)
                {
                    Paragraph paragraph = element as Paragraph;
                    var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
                    string elementId = paragraph.ParagraphId ?? WordDocumentPartAttributes.GetParagraphNoIdFormatter(_paragraphCounter);

                    if (numberingProperties != null)
                    {
                        indent += numberingProperties.NumberingLevelReference.Val.Value;
                        if (WordDocumentPartAttributes.BulletListIds.Any(b => b == numberingProperties.NumberingId.Val?.Value))
                            elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.BulletList);
                        else if (WordDocumentPartAttributes.NumberedListIds.Any(b => b == numberingProperties.NumberingId.Val?.Value))
                            elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.NumberedList);
                        else
                            elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.BulletList);
                    }
                    else
                        elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), indent, Helpers.ElementType.Paragraph);

                    _paragraphCounter++;
                }
                else if (element is Picture || element is Drawing)
                {
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), element.LocalName, indent, Helpers.ElementType.Picture);
                }
                else if (element is Table)
                {
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), (element as Table).LocalName, indent, Helpers.ElementType.Table);
                }
                else
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), WordTools.GetElementName(element, WordDocumentPartAttributes.MaxNameLength), indent);

                result.Add(elementToAdd);
                if (element.HasChildren)
                {
                    foreach (var elmentChild in element.ChildElements)
                    {
                        _index++;
                        CreatePartsSelectionTreeElements(elmentChild, _index, supportedType, ++indent);
                    }

                }
            }

            return result;
        }
    }
}
