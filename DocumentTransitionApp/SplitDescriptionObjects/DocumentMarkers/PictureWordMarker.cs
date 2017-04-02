using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using OpenXMLTools.Word.OpenXmlElements;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects.DocumentMarkers
{
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

        public static PersonPictureMarker[] GetPictureMarkers(IEnumerable<PartsSelectionTreeElement> parts)
        {
            IList<PersonPictureMarker> result = new List<PersonPictureMarker>();

            foreach (var part in parts.Where(p => p.Type == ElementType.Picture))
            {
                if (part.Parent == null || (part.Parent != null && !part.Parent.Selected))
                {
                    PersonPictureMarker pictureMarker = new PersonPictureMarker();
                    pictureMarker.ElementId = part.ElementId;
                    result.Add(pictureMarker);
                }
            }

            return result.ToArray();
        }
    }
}
