using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

namespace SplitDescriptionObjects
{
    public interface IExcelMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedElements(string id, string id2);
    }

    public abstract class ExcelMarker : IExcelMarker
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

    public interface IUniversalExcelMarker : IExcelMarker
    {
    }

    public class UniversalExcelMarker : ExcelMarker, IUniversalExcelMarker
    {
        public UniversalExcelMarker(Workbook body) :
            base(body)
        {
        }
    }
}
