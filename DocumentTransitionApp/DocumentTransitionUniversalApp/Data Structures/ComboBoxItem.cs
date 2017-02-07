using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentTransitionUniversalApp.Data_Structures
{
    public class ComboBoxItem
    {
        public string Name { get; set; }
        public int Id { get; set; }

        public static ComboBoxItem GetComboBoxItemByName(IEnumerable<ComboBoxItem> items, string name)
        {
            return items.Single(it => it.Name == name);
        }

        public static ComboBoxItem GetComboBoxItemById(IEnumerable<ComboBoxItem> items, int id)
        {
            return items.Single(it => it.Id == id);
        }
    }
}
