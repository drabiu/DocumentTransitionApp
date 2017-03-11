using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OpenXMLTools.Interfaces
{
    public interface IExcelTools
    {
        GetMissingSharedStringItemsResult GetMissingSharedStringItems(SpreadsheetDocument target, SpreadsheetDocument source);
        SpreadsheetDocument MergeWorkSheets(SpreadsheetDocument target, SpreadsheetDocument source);
    }
}
