using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects.DocumentMarkers
{
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
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, ElementsList, el => el is Paragraph || !WordDocumentPartAttributes.IsSupportedType(el), element => GetParagraphId(element));

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

        public static PersonUniversalMarker[] GetUniversalMarkers(IEnumerable<PartsSelectionTreeElement> parts)
        {
            IList<PersonUniversalMarker> result = new List<PersonUniversalMarker>();
            var paragraphParts = parts.Where(p => p.Type == ElementType.Paragraph).ToList();
            int lastPartId = -2;
            int iterationCounter = paragraphParts.Count;
            PersonUniversalMarker universalMarker = new PersonUniversalMarker();
            foreach (var part in paragraphParts.OrderBy(p => p.Id))
            {
                iterationCounter--;
                if (lastPartId + 1 == int.Parse(part.Id))
                {
                    universalMarker.SelectionLastelementId = part.ElementId;
                }
                else
                {
                    if (lastPartId > 0)
                        result.Add(universalMarker);

                    universalMarker = new PersonUniversalMarker();
                    universalMarker.ElementId = part.ElementId;
                    universalMarker.SelectionLastelementId = part.ElementId;
                }

                if (iterationCounter == 0)
                    result.Add(universalMarker);

                lastPartId = int.Parse(part.Id);
            }

            return result.ToArray();
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
}
