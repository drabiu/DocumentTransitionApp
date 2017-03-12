using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OpenXMLTools
{
    public class SharedStringItemsComparer : IEqualityComparer<SharedStringItem>
    {
        public bool Equals(SharedStringItem x, SharedStringItem y)
        {
            return string.Equals(x.InnerText, y.InnerText);
        }

        public int GetHashCode(SharedStringItem obj)
        {
            return obj.InnerText.GetHashCode();
        }
    }
}
