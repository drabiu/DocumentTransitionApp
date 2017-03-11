using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace OpenXMLTools
{
    public class SharedStringIndex
    {
        public int OldIndex { get; set; }
        public int NewIndex { get; set; }

        public SharedStringIndex(int oldIndex, int newIndex)
        {
            OldIndex = oldIndex;
            NewIndex = newIndex;
        }
    }

    public class GetMissingSharedStringItemsResult
    {
        public IList<SharedStringItem> SharedStringItems { get; set; }
        public IList<SharedStringIndex> SharedStringIndexes { get; set; }

        public GetMissingSharedStringItemsResult()
        {
            SharedStringItems = new List<SharedStringItem>();
            SharedStringIndexes = new List<SharedStringIndex>();
        }
    }
}
