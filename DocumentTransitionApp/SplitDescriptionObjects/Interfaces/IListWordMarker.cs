using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface IListWordMarker : IWordMarker
    {
        List<MarkerWordSelector> GetSubdividedParts(Person person);
    }
}
