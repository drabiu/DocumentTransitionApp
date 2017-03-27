using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface ITableWordMarker : IWordMarker
    {
        List<MarkerWordSelector> GetSubdividedParts(Person person);
    }
}
