using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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

    public class TableWordMarker : WordMarker, ITableWordMarker
    {
        public TableWordMarker(Body body) :
            base(body)
        {

        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
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
    }

    public class ListWordMarker : WordMarker, IListWordMarker
    {
        public ListWordMarker(Body body) :
            base(body)
        {

        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
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
    }

    public class PictureWordMarker : WordMarker, IPictureWordMarker
    {
        public PictureWordMarker(Body body) : base(body)
        {

        }

        public static void SetPartsOwner(List<PartsSelectionTreeElement> parts, Person person)
        {
            foreach (var pictureMarker in person.PictureMarker)
            {
                var selectedPartsIndexes = MarkerHelper<PartsSelectionTreeElement>.GetCrossedElements(pictureMarker.ElementId, pictureMarker.ElementId, parts, element => element.ElementId);
                foreach (var index in selectedPartsIndexes)
                {
                    parts[index].OwnerName = person.Email;
                    parts[index].Selected = true;
                }
            }
        }
    }
}
