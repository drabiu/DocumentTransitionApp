using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OpenXMLTools.Interfaces
{
    public interface IExcelTools
    {
        IEnumerable<SharedStringItem> GetAddedSharedStringItems(SpreadsheetDocument target);
    }
}
