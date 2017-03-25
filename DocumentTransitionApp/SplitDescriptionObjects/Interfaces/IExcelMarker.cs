using System.Collections.Generic;

namespace SplitDescriptionObjects.Interfaces
{
    public interface IExcelMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedSheetElements(string id, string id2);
    }
}
