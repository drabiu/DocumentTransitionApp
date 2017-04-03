using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools;
using OpenXMLTools.Helpers;
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
        public const string CounterName = "counter";
        public const string NumIdTag = "[numId]";

        public static string GetParagraphNoIdFormatter(int id)
        {
            return string.Format("{0}{1}", ParagraphHasNoId, id);
        }

        public static string GetParagraphListIdFormatter(string id, int numberingId)
        {
            return string.Format("{0}{1}{2}", id, NumIdTag, numberingId);
        }

        public static string GetParagraphIdFromListIdFormatter(string paragraphListId)
        {
            return paragraphListId.Split(new string[] { NumIdTag }, StringSplitOptions.None).First();
        }

        public static int GetNumberingIdFromListId(string paragraphListId)
        {
            return int.Parse(paragraphListId.Split(new string[] { NumIdTag }, StringSplitOptions.None).Last());
        }

        public static string GetTableIdFormatter(int id)
        {
            return string.Format("tab{0}", id);
        }

        public static string GetDrawingIdFormatter(int id)
        {
            return string.Format("drw{0}", id);
        }

        public static bool IsSupportedType(OpenXmlElement element)
        {
            bool isSupported = false;
            isSupported = element is Paragraph
                //|| element is Picture
                || element is Table
                || element is Drawing
                || element is Run;

            return isSupported;
        }
    }

    public class WordDocumentParts : IDocumentParts
    {
        NameIndexer _indexer;
        DocumentParts _documentParts;

        public WordDocumentParts()
        {
            _documentParts = new DocumentParts(this);
            _indexer = new NameIndexer(new List<string>() { WordDocumentPartAttributes.CounterName });
        }

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
                        bool visible = true;
                        //group list elements to one
                        if (siblingsList.Contains(element))
                        {
                            visible = false;
                        }
                        else if (element is Paragraph && WordTools.IsListParagraph(element as Paragraph))
                        {
                            siblingsList = WordTools.GetAllSiblingListElements(element as Paragraph, body.ChildElements.ToList(), WordTools.GetNumberingId(element as Paragraph));
                        }

                        documentElements.AddRange(_documentParts.CreatePartsSelectionTreeElements(element, null, _documentParts.Index, supportedType, 0, visible));
                    }
                }
            }

            return documentElements;
        }

        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            return Get(file, el => WordDocumentPartAttributes.IsSupportedType(el));
        }

        public PartsSelectionTreeElement GetParagraphSelectionTreeElement(OpenXmlElement element, PartsSelectionTreeElement parent, ref int id, Predicate<OpenXmlElement> supportedType, int indent, bool visible)
        {
            PartsSelectionTreeElement elementToAdd = null;
            if (element is Paragraph)
            {
                ParagraphDecorator paragraphDecorator = new ParagraphDecorator(element);
                string paragraphId = paragraphDecorator.GetParagraph().ParagraphId ?? WordDocumentPartAttributes.GetParagraphNoIdFormatter(_indexer.GetNextIndex(WordDocumentPartAttributes.CounterName, paragraphDecorator.GetElementType()));
                string elementId = string.Empty;
                if (paragraphDecorator.GetElementType() == ElementType.NumberedList || paragraphDecorator.GetElementType() == ElementType.BulletList)
                    elementId = WordDocumentPartAttributes.GetParagraphListIdFormatter(paragraphId, WordTools.GetNumberingId(paragraphDecorator.GetParagraph()));
                else
                    elementId = paragraphId;

                elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, paragraphDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, paragraphDecorator.GetElementType());
                elementToAdd.Visible = visible;
                id++;
            }
            else if (element is Drawing)
            {
                DrawingDecorator drawingDecorator = new DrawingDecorator(element);
                string elementId = WordDocumentPartAttributes.GetDrawingIdFormatter(_indexer.GetNextIndex(WordDocumentPartAttributes.CounterName, drawingDecorator.GetElementType()));
                elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, drawingDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, drawingDecorator.GetElementType());
                id++;
            }
            //else if (element is Picture)
            //{

            //}
            else if (element is Table)
            {
                TableDecorator tableDecorator = new TableDecorator(element);
                string elementId = WordDocumentPartAttributes.GetTableIdFormatter(_indexer.GetNextIndex(WordDocumentPartAttributes.CounterName, tableDecorator.GetElementType()));
                elementToAdd = new PartsSelectionTreeElement(id.ToString(), elementId, tableDecorator.GetElementName(WordDocumentPartAttributes.MaxNameLength), indent, tableDecorator.GetElementType());
                id++;
            }

            return elementToAdd;
        }
    }
}
