using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTools.Interfaces
{
    public interface IExcelTools
    {
        GetMergedSharedStringItemsResult GetMergedSharedStringItems(SpreadsheetDocument target, SpreadsheetDocument source);
        SpreadsheetDocument MergeWorkSheets(SpreadsheetDocument target, SpreadsheetDocument source);
    }
}
