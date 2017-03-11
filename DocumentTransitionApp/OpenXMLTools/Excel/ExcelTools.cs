using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLTools.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXMLTools
{
    public class ExcelTools : IExcelTools
    {
        #region Public methods

        public SpreadsheetDocument MergeWorkSheets(SpreadsheetDocument target, SpreadsheetDocument source)
        {
            ExcelMergeData mergeData = new ExcelMergeData();
            mergeData.SetDocumentData(source);

            var missingSharedStringItems = GetMissingSharedStringItems(target, source);
            target.WorkbookPart.SharedStringTablePart.SharedStringTable.Append(missingSharedStringItems.SharedStringItems);
            FixSharedStringReference(mergeData.WorksheetPartList, missingSharedStringItems.SharedStringIndexes);

            //delete all related worksheetparts
            foreach (Sheet element in mergeData.Sheets)
            {
                //target.WorkbookPart.DeletePart(element.Id);
            }

            //check if parts relationship id won`t repeat
            foreach (KeyValuePair<string, WorksheetPart> element in mergeData.WorksheetPartList)
            {
                //var elementId = workbook.GetIdOfPart(element);
                //target.WorkbookPart.AddPart(element.Value, element.Key);
            }

            target.WorkbookPart.Workbook.Sheets.Append(mergeData.Sheets);
            //target.WorkbookPart.Workbook.Save();
            target.Save();

            return target;
        }

        public GetMissingSharedStringItemsResult GetMissingSharedStringItems(SpreadsheetDocument target, SpreadsheetDocument source)
        {
            var targetItems = target.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Select(el => el.CloneNode(true) as SharedStringItem).ToList();
            var targetItemsText = new HashSet<string>(targetItems.Select(t => t.Text.Text));
            var sourceItems = source.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Select(el => el.CloneNode(true) as SharedStringItem);

            GetMissingSharedStringItemsResult result = new GetMissingSharedStringItemsResult();
            int sourceItemIndex = 0;
            foreach (SharedStringItem item in sourceItems)
            {
                if (!targetItemsText.Contains(item.Text.Text))
                {
                    result.SharedStringItems.Add(item);
                    result.SharedStringIndexes.Add(new SharedStringIndex(sourceItemIndex, targetItems.Count - 1));
                }

                sourceItemIndex++;
            }

            return result;
        }

        #endregion

        #region Private methods

        private void RemoveReferencesFromCalculationChainPart()
        {
            //when spliting need to remove unused references to removed sheets
            //excelDoc.WorkbookPart.CalculationChainPart 
        }
      
        private IDictionary<string, WorksheetPart> FixSharedStringReference(IDictionary<string, WorksheetPart> workSheetPartList, IList<SharedStringIndex> indexes)
        {
            var oldIndexes = new HashSet<string>(indexes.Select(i => i.OldIndex.ToString()));
            foreach (KeyValuePair<string, WorksheetPart> element in workSheetPartList)
            {
                var cells = element.Value.Worksheet.Descendants<Cell>().Where(cell => cell?.DataType?.Value == CellValues.SharedString);
                foreach (Cell cell in cells)
                {
                    if (oldIndexes.Contains(cell.CellValue.Text))
                    {
                        int cellOldIndex = int.Parse(cell.CellValue.Text);
                        cell.CellValue.Text = indexes.First(i => i.OldIndex == cellOldIndex).NewIndex.ToString();
                    }
                }              
            }

            return workSheetPartList;
        }

        #endregion  

    }
}
