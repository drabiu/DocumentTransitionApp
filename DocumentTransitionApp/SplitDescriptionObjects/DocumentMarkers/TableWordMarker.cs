using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects.DocumentMarkers
{
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

        public static PersonTableMarker[] GetTableMarkers(IEnumerable<PartsSelectionTreeElement> parts)
        {
            IList<PersonTableMarker> result = new List<PersonTableMarker>();

            foreach (var part in parts.Where(p => p.Type == ElementType.Table))
            {
                PersonTableMarker tableMarker = new PersonTableMarker();
                tableMarker.ElementId = part.ElementId;
                result.Add(tableMarker);
            }

            return result.ToArray();
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
}
