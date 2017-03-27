using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface IPictureWordMarker : IWordMarker
    {
        List<MarkerWordSelector> GetSubdividedParts(Person person);
    }
}
