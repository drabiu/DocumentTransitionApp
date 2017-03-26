using DocumentTransitionUniversalApp.Views;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DocumentTransitionUniversalApp.Data_Structures
{
    public class WordPartsPageData
    {
        public List<ComboBoxItem> ComboItems { get; set; }
        public int LastId { get; set; }
        public static int AllItemsId = 0;
        public ObservableCollection<PartsSelectionTreeElement> SelectionParts { get; set; }

        public WordPartsPageData()
        {
            ComboItems = new List<ComboBoxItem>();
            ComboItems.Add(new ComboBoxItem() { Id = LastId = AllItemsId, Name = "All" });
            SelectionParts = new ObservableCollection<PartsSelectionTreeElement>();
        }

        public WordPartsPageData(WordSelectPartsPage page)
        {
            ComboItems = page._pageData.ComboItems;
            LastId = page._pageData.LastId;
            SelectionParts = page._pageData.SelectionParts;

            page.CopyDataToControl(this);
        }
    }
}
