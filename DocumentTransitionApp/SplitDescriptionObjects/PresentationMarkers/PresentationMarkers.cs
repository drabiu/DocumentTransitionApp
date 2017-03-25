using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using SplitDescriptionObjects.Interfaces;

namespace SplitDescriptionObjects
{
    public abstract class PresentationMarker : IPresentationMarker
    {
        PresentationPart DocumentBody;

        public PresentationMarker(PresentationPart body)
        {
            DocumentBody = body;
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        public IList<int> GetCrossedSlideIdElements(string id, string id2)
        {
            var elements = DocumentBody.Presentation.SlideIdList.Elements<SlideId>();
            var indexes = MarkerHelper<SlideId>.GetCrossedElements(id, id2, elements.ToList(), element => element.RelationshipId);

            return indexes;
        }
    }

    public class SlidePresentationMarker : PresentationMarker, ISlidePresentationMarker
    {
        public SlidePresentationMarker(PresentationPart body) :
            base(body)
        {
        }
    }
}
