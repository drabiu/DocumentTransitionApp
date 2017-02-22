using System;
using System.Collections.Generic;
using System.Linq;

using Present = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace SplitDescriptionObjects
{
    public interface IPresentationMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedSlideIdElements(string id, string id2);
    }

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
            var elements = DocumentBody.Presentation.SlideIdList.Elements<Present.SlideId>();
            var indexes = MarkerHelper<Present.SlideId>.GetCrossedElements(id, id2, elements.ToList(), element => element.RelationshipId);

            return indexes;
        }
    }

    public interface IUniversalPresentationMarker : IPresentationMarker
    {
    }

    public class UniversalPresentationMarker : PresentationMarker, IUniversalPresentationMarker
    {
        public UniversalPresentationMarker(PresentationPart body) :
            base(body)
        {
        }
    }
}
