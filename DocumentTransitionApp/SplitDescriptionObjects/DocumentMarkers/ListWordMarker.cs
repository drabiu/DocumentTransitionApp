using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects.DocumentMarkers
{
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

        public static PersonListMarker[] GetListMarkers(IEnumerable<PartsSelectionTreeElement> parts)
        {
            IList<PersonListMarker> result = new List<PersonListMarker>();
            var visibleParts = parts.Where(p => p.Visible);
            foreach (var part in visibleParts.Where(p => p.Type == ElementType.BulletList || p.Type == ElementType.NumberedList))
            {
                PersonListMarker listMarker = new PersonListMarker();
                listMarker.ElementId = part.ElementId;
                var siblings = GetSiblings(parts.ToList(), part);
                if (siblings.Count > 0)
                {
                    var lastSibling = siblings.Last();
                    listMarker.SelectionLastelementId = lastSibling.ElementId;

                }
                else
                {
                    listMarker.SelectionLastelementId = part.ElementId;
                }

                result.Add(listMarker);
            }

            return result.ToArray();
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
}
