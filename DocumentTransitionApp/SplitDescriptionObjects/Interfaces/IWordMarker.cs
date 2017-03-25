using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface IWordMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedParagraphElements(string id, string id2);
    }
}
