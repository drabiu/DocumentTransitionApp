using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using OpenXMLTools.Word.OpenXmlElements;
using SplitDescriptionObjects.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects
{
    public abstract class WordMarker : IWordMarker
    {
        protected Body DocumentBody;
        protected List<OpenXmlElement> ElementsList;

        public WordMarker(Body body)
        {
            DocumentBody = body;
            ElementsList = DocumentBody.ChildElements.ToList();
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        protected static void SelectChildParts(IEnumerable<PartsSelectionTreeElement> parts, Person person)
        {
            foreach (var part in parts)
            {
                foreach (var child in part.Childs)
                {
                    SelectChildParts(child.Childs, person);
                    child.OwnerName = person.Email;
                    child.Selected = true;
                }
            }
        }
    }

    public class UniversalWordMarker : WordMarker, IUniversalWordMarker
    {
        List<MarkerWordSelector> _subdividedParagraphs;
        List<OpenXmlElement> ParagraphsList;

        public UniversalWordMarker(Body body, List<MarkerWordSelector> subdividedParagraphs) :
            base(body)
        {
            _subdividedParagraphs = subdividedParagraphs;
            ElementsList = _subdividedParagraphs.Select(sp => sp.Element).ToList();
            ParagraphsList = ElementsList.Where(el => el is Paragraph).ToList();
        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
        {
            SelectCrossedUniversalParts(parts, person);
            SelectChildParts(parts, person);
        }

        public IList<int> GetCrossedParagraphElements(string id, string id2)
        {
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, ElementsList, el => el is Paragraph, element => GetParagraphId(element));

            return indexes;
        }

        public List<MarkerWordSelector> GetSubdividedParts(Person person)
        {
            if (person.UniversalMarker != null)
            {
                foreach (PersonUniversalMarker marker in person.UniversalMarker)
                {
                    IList<int> result = GetCrossedParagraphElements(marker.ElementId, marker.SelectionLastelementId);
                    foreach (int index in result)
                    {
                        if (string.IsNullOrEmpty(_subdividedParagraphs[index].Email))
                        {
                            _subdividedParagraphs[index].Email = person.Email;
                        }
                    }
                }
            }

            return _subdividedParagraphs;
        }

        private static void SelectCrossedUniversalParts(List<PartsSelectionTreeElement> parts, Person person)
        {
            if (person.UniversalMarker != null)
            {
                foreach (var universalMarker in person.UniversalMarker)
                {
                    var selectedPartsIndexes = MarkerHelper<PartsSelectionTreeElement>.GetCrossedElements(universalMarker.ElementId, universalMarker.SelectionLastelementId, parts, element => element.ElementId);
                    foreach (var index in selectedPartsIndexes)
                    {
                        parts[index].OwnerName = person.Email;
                        parts[index].Selected = true;
                    }
                }
            }
        }

        private string GetParagraphId(OpenXmlElement element)
        {
            string result = string.Empty;
            if (element is Paragraph)
            {
                Paragraph parahraph = (element as Paragraph);
                int index = ParagraphsList.FindIndex(el => el.Equals(element));
                result = parahraph.ParagraphId != null ? parahraph.ParagraphId.Value : WordDocumentPartAttributes.GetParagraphNoIdFormatter(index);
            }

            return result;
        }
    }

    public class TableWordMarker : WordMarker, ITableWordMarker
    {
        List<MarkerWordSelector> _subdividedParagraphs;
        List<OpenXmlElement> TablesList;

        public TableWordMarker(Body body, List<MarkerWordSelector> subdividedParagraphs) :
            base(body)
        {
            _subdividedParagraphs = subdividedParagraphs;
            ElementsList = _subdividedParagraphs.Select(sp => sp.Element).ToList();
            TablesList = ElementsList.Where(sp => sp is Table).ToList();
        }

        public IList<int> GetCrossedTableElements(string id, string id2)
        {
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, ElementsList, el => el is Table, element => GetTableId(element));

            return indexes;
        }

        public List<MarkerWordSelector> GetSubdividedParts(Person person)
        {
            if (person.TableMarker != null)
            {
                foreach (PersonTableMarker marker in person.TableMarker)
                {
                    IList<int> result = GetCrossedTableElements(marker.ElementId, marker.ElementId);
                    foreach (int index in result)
                    {
                        if (string.IsNullOrEmpty(_subdividedParagraphs[index].Email))
                        {
                            _subdividedParagraphs[index].Email = person.Email;
                        }
                    }
                }
            }

            return _subdividedParagraphs;
        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
        {
            if (person.TableMarker != null)
            {
                foreach (var tableMarker in person.TableMarker)
                {
                    var selectedPartsIndexes = MarkerHelper<PartsSelectionTreeElement>.GetCrossedElements(tableMarker.ElementId, tableMarker.ElementId, parts, element => element.ElementId);
                    foreach (var index in selectedPartsIndexes)
                    {
                        parts[index].OwnerName = person.Email;
                        parts[index].Selected = true;
                    }
                }
            }

            //SelectChildParts(parts, person);
        }

        private string GetTableId(OpenXmlElement element)
        {
            string result = string.Empty;
            if (element is Table)
            {
                int index = TablesList.FindIndex(el => el.Equals(element));
                result = WordDocumentPartAttributes.GetTableIdFormatter(index);
            }

            return result;
        }
    }

    public class ListWordMarker : WordMarker, IListWordMarker
    {
        List<MarkerWordSelector> _subdividedParagraphs;
        List<OpenXmlElement> ParagraphsList;

        public ListWordMarker(Body body, List<MarkerWordSelector> subdividedParagraphs) :
            base(body)
        {
            _subdividedParagraphs = subdividedParagraphs;
            ElementsList = _subdividedParagraphs.Select(sp => sp.Element).ToList();
            ParagraphsList = ElementsList.Where(el => el is Paragraph).ToList();
        }

        public IList<int> GetCrossedListElements(string id, string id2)
        {
            id = WordDocumentPartAttributes.GetParagraphIdFromListIdFormatter(id);
            id2 = WordDocumentPartAttributes.GetParagraphIdFromListIdFormatter(id2);
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, ElementsList, el => el is Paragraph, element => GetListId(element));

            return indexes;
        }

        public List<MarkerWordSelector> GetSubdividedParts(Person person)
        {
            if (person.ListMarker != null)
            {
                foreach (PersonListMarker marker in person.ListMarker)
                {
                    IList<int> result = GetCrossedListElements(marker.ElementId, marker.SelectionLastelementId);
                    foreach (int index in result)
                    {
                        if (string.IsNullOrEmpty(_subdividedParagraphs[index].Email))
                        {
                            _subdividedParagraphs[index].Email = person.Email;
                        }
                    }
                }
            }

            return _subdividedParagraphs;
        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
        {
            if (person.ListMarker != null)
            {
                foreach (var listMarker in person.ListMarker)
                {
                    var selectedPartsIndexes = MarkerHelper<PartsSelectionTreeElement>.GetCrossedElements(listMarker.ElementId, listMarker.SelectionLastelementId, parts, element => element.ElementId);
                    foreach (var index in selectedPartsIndexes)
                    {
                        parts[index].OwnerName = person.Email;
                        parts[index].Selected = true;
                    }
                }
            }

            //SelectChildParts(parts, person);
        }

        public static IList<PartsSelectionTreeElement> GetSiblings(List<PartsSelectionTreeElement> parts, PartsSelectionTreeElement part)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            var index = parts.FindIndex(e => e.ElementId == part.ElementId);
            if (part.IsListElement())
            {
                var partNumberingId = WordDocumentPartAttributes.GetNumberingIdFromListId(part.ElementId);
                foreach (var element in parts.Skip(index + 1))
                {
                    if (element.IsListElement() && WordDocumentPartAttributes.GetNumberingIdFromListId(element.ElementId) == partNumberingId)
                        result.Add(element);
                    else
                        break;
                }
            }

            return result;
        }

        private string GetListId(OpenXmlElement element)
        {
            string result = string.Empty;
            if (element is Paragraph)
            {
                Paragraph parahraph = (element as Paragraph);
                result = parahraph.ParagraphId.Value;
            }

            return result;
        }
    }

    public class PictureWordMarker : WordMarker, IPictureWordMarker
    {
        List<MarkerWordSelector> _subdividedParagraphs;

        public PictureWordMarker(Body body, List<MarkerWordSelector> subdividedParagraphs) : base(body)
        {
            _subdividedParagraphs = subdividedParagraphs;
            ElementsList = _subdividedParagraphs.Select(sp => sp.Element).ToList();
        }

        public static void SetPartsOwner(IList<PartsSelectionTreeElement> parts, Person person)
        {
            var pictureParts = parts.Where(p => p.Type == ElementType.Picture);
            foreach (var part in parts)
                SetPartsOwner(part.Childs, person);

            foreach (var picturePart in pictureParts)
                SelectPicturePart(picturePart, person);
        }

        private static void SelectPicturePart(PartsSelectionTreeElement part, Person person)
        {
            if (person.PictureMarker != null)
            {
                foreach (var pictureMarker in person.PictureMarker)
                {
                    if (part.ElementId == pictureMarker.ElementId)
                    {
                        part.OwnerName = person.Email;
                        part.Selected = true;
                    }
                }
            }
        }

        public List<MarkerWordSelector> GetSubdividedParts(Person person)
        {
            if (person.PictureMarker != null)
            {
                foreach (PersonPictureMarker marker in person.PictureMarker)
                {
                    int drawingIndex = 0;
                    for (int index = 0; index < ElementsList.Count; index++)
                    {
                        var element = ElementsList[index];
                        var drawings = element.Descendants<Drawing>();
                        foreach (var drawing in drawings)
                        {
                            if (WordDocumentPartAttributes.GetDrawingIdFormatter(drawingIndex) == marker.ElementId)
                            {
                                DrawingDecorator drawingDecorator = new DrawingDecorator(drawing);
                                MarkerWordSelector markerWordSelector = new MarkerWordSelector(drawingDecorator.CreateParagraph(), person.Email);
                                _subdividedParagraphs.Insert(index, markerWordSelector);
                            }

                            drawingIndex++;
                        }
                    }
                }
            }

            return _subdividedParagraphs;
        }
    }
}
