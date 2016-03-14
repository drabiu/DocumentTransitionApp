using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

namespace SplitDescriptionObjects
{
    public abstract class ExcelMarker : IDocumentMarker
    {
        Workbook DocumentBody;

        public ExcelMarker(Workbook body)
        {
            DocumentBody = body;
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        public IList<int> GetCrossedElements(string id, string id2)
        {
            throw new NotImplementedException();
        }
    }
}
