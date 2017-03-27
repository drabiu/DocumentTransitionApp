using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface IUniversalWordMarker : IWordMarker
    {
        List<MarkerWordSelector> GetSubdividedParts(Person person);
        IList<int> GetCrossedParagraphElements(string id, string id2);
    }
}
