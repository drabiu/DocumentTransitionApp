using DocumentEditPartsEngine;
using System.Collections.Generic;

namespace DocumentSplitEngineTests.Mocks
{
    public class PartsSelectionTreeElementMock
    {
        public static IList<PartsSelectionTreeElement> GetListMock()
        {
            IList<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            result.Add(new PartsSelectionTreeElement("1", "el1", "name1", 0) { OwnerName = "test1", Selected = true });
            result.Add(new PartsSelectionTreeElement("2", "el2", "name2", 0) { OwnerName = "test1", Selected = true });
            result.Add(new PartsSelectionTreeElement("3", "el3", "name3", 0) { OwnerName = "test1", Selected = true });

            result.Add(new PartsSelectionTreeElement("4", "el4", "name4", 0));
            result.Add(new PartsSelectionTreeElement("5", "el5", "name5", 0) { OwnerName = "test2", Selected = true });
            result.Add(new PartsSelectionTreeElement("6", "el6", "name6", 0));
            result.Add(new PartsSelectionTreeElement("7", "el7", "name7", 0) { OwnerName = "test2", Selected = true });

            return result;
        }

        public static List<PartsSelectionTreeElement> GetUnselectedPartsListMock()
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            result.Add(new PartsSelectionTreeElement("1", "el1", "name1", 0));
            result.Add(new PartsSelectionTreeElement("2", "el2", "name2", 0));
            result.Add(new PartsSelectionTreeElement("3", "el3", "name3", 0));

            result.Add(new PartsSelectionTreeElement("4", "el4", "name4", 0));
            result.Add(new PartsSelectionTreeElement("5", "el5", "name5", 0));
            result.Add(new PartsSelectionTreeElement("6", "el6", "name6", 0));
            result.Add(new PartsSelectionTreeElement("7", "el7", "name7", 0));

            return result;
        }
    }
}
