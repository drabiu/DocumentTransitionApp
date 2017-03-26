using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using SplitDescriptionObjects.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects
{
    public abstract class WordMarker : IWordMarker
    {
        Body DocumentBody;
        List<OpenXmlElement> ElementsList;

        public WordMarker(Body body)
        {
            DocumentBody = body;
            ElementsList = DocumentBody.ChildElements.ToList();
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        public IList<int> GetCrossedParagraphElements(string id, string id2)
        {
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, ElementsList, el => el is Paragraph, element => GetParagraphId(element));

            return indexes;
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

        private string GetParagraphId(OpenXmlElement element)
        {
            string result = string.Empty;
            if (element is Paragraph)
            {
                Paragraph parahraph = (element as Paragraph);
                int index = ElementsList.FindIndex(el => el.Equals(element));
                result = parahraph.ParagraphId != null ? parahraph.ParagraphId.Value : WordDocumentPartAttributes.GetParagraphNoIdFormatter(index);
            }

            return result;
        }
    }

    public class UniversalWordMarker : WordMarker, IUniversalWordMarker
    {
        public UniversalWordMarker(Body body) :
            base(body)
        {
        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
        {
            SelectCrossedUniversalParts(parts, person);
            SelectChildParts(parts, person);
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
    }

    public class TableWordMarker : WordMarker, ITableWordMarker
    {
        public TableWordMarker(Body body) :
            base(body)
        {

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

            SelectChildParts(parts, person);
        }
    }

    public class ListWordMarker : WordMarker, IListWordMarker
    {
        public ListWordMarker(Body body) :
            base(body)
        {

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

            SelectChildParts(parts, person);
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
    }

    public class PictureWordMarker : WordMarker, IPictureWordMarker
    {
        public PictureWordMarker(Body body) : base(body)
        {

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
    }
}
