using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface IPresentationMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedSlideIdElements(string id, string id2);
    }
}
