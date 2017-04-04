using DocumentEditPartsEngine;
using OpenXMLTools.Helpers;
using System.Collections.Generic;

namespace SplitDescriptionObjectsTests.Mocks
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
            result.Add(new PartsSelectionTreeElement("8", "el8", "name8", 0, ElementType.Table) { OwnerName = "test2", Selected = true });
            result.Add(new PartsSelectionTreeElement("9", "el9", "name9", 0, ElementType.Table));
            result.Add(new PartsSelectionTreeElement("10", "el10", "name10", 0, ElementType.Table) { OwnerName = "test1", Selected = true });
            result.Add(new PartsSelectionTreeElement("11", "el11", "name11", 0, ElementType.Table) { OwnerName = "test1", Selected = true });
            result.Add(new PartsSelectionTreeElement("12", "el12[numId]2", "name12", 0, ElementType.BulletList) { OwnerName = "test1", Selected = true });
            result.Add(new PartsSelectionTreeElement("13", "el13[numId]2", "name13", 0, ElementType.NumberedList) { OwnerName = "test2", Selected = true });
            result.Add(new PartsSelectionTreeElement("14", "el14[numId]2", "name14", 0, ElementType.NumberedList) { OwnerName = "test2", Selected = true, Visible = false });
            result.Add(new PartsSelectionTreeElement("15", "el15", "name15", 0, ElementType.Picture) { OwnerName = "test1", Selected = true });
            result.Add(new PartsSelectionTreeElement("16", "el16", "name16", 0, ElementType.Picture) { OwnerName = "test2", Selected = true });

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
