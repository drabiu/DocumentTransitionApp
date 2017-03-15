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

            var mergedSharedStringItemsResult = GetMergedSharedStringItems(target, source);
            target.WorkbookPart.SharedStringTablePart.SharedStringTable.RemoveAllChildren();
            target.WorkbookPart.SharedStringTablePart.SharedStringTable.Append(mergedSharedStringItemsResult.SharedStringItems);
            FixSharedStringReference(mergeData.WorksheetPartList, mergedSharedStringItemsResult.SharedStringIndexes);

            DeleteSheetsAndReferencedWorksheetParts(target, mergeData);
            ReplaceWorkSheetparts(target, mergeData);

            FixSheetsIds(target, mergeData);
            GetMergedCalculationChainPart();
            target.WorkbookPart.Workbook.Sheets.Append(mergeData.Sheets);

            target.Save();

            return target;
        }

        public GetMergedSharedStringItemsResult GetMergedSharedStringItems(SpreadsheetDocument target, SpreadsheetDocument source)
        {
            var targetItems = target.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();
            var sourceItems = source.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();

            var mergedItems = targetItems.Union(sourceItems, new SharedStringItemsComparer()).ToArray();

            GetMergedSharedStringItemsResult result = new GetMergedSharedStringItemsResult();
            result.SharedStringItems.AddRange(mergedItems.Select(el => el.CloneNode(true) as SharedStringItem));

            int sourceItemIndex = 0;
            foreach (SharedStringItem item in sourceItems)
            {
                int newIndex = Array.FindIndex(mergedItems, m => m.InnerText == item.InnerText);
                if (newIndex != -1 && newIndex != sourceItemIndex)
                {
                    result.SharedStringIndexes.Add(new SharedStringIndex(sourceItemIndex, newIndex));
                }

                sourceItemIndex++;
            }

            return result;
        }

        public SpreadsheetDocument RemoveReferencedCalculationChainCell(SpreadsheetDocument target)
        {
            var calculationCells = target.WorkbookPart.CalculationChainPart.CalculationChain.Elements<CalculationCell>().ToList();
            var sheetIds = new HashSet<uint>(target.WorkbookPart.Workbook.Sheets.Select(s => (s as Sheet).SheetId.Value));
            foreach (var calculationCell in calculationCells)
            {
                if (!sheetIds.Contains((uint)calculationCell.SheetId.Value))
                    target.WorkbookPart.CalculationChainPart.CalculationChain.RemoveChild(calculationCell);
            }

            target.WorkbookPart.CalculationChainPart.CalculationChain.Save();

            var calculationChainPart = target.WorkbookPart.CalculationChainPart;
            if (calculationChainPart.CalculationChain.ChildElements.Count == 0)
                target.WorkbookPart.DeletePart(calculationChainPart);

            return target;
        }

        #endregion

        #region Static public methods

        public static string GetRowTitle(Row row, int nameLength)
        {
            return string.Empty;
        }

        #endregion

        #region Private methods

        private void GetMergedCalculationChainPart()
        {
            //when spliting need to remove unused references to removed sheets
            //excelDoc.WorkbookPart.CalculationChainPart 
        }

        private Dictionary<string, WorksheetPart> FixSharedStringReference(Dictionary<string, WorksheetPart> workSheetPartList, IList<SharedStringIndex> indexes)
        {
            var oldIndexes = new HashSet<string>(indexes.Select(i => i.OldIndex.ToString()));
            foreach (KeyValuePair<string, WorksheetPart> element in workSheetPartList)
            {
                var cells = element.Value.Worksheet.Descendants<Cell>().Where(cell => cell != null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString);
                foreach (Cell cell in cells)
                {
                    if (oldIndexes.Contains(cell.CellValue.InnerText))
                    {
                        int cellOldIndex = int.Parse(cell.CellValue.InnerText);
                        cell.CellValue.Text = indexes.First(i => i.OldIndex == cellOldIndex).NewIndex.ToString();
                        element.Value.Worksheet.Save();
                    }
                }
            }

            return workSheetPartList;
        }

        private SpreadsheetDocument DeleteSheetsAndReferencedWorksheetParts(SpreadsheetDocument target, ExcelMergeData mergeData)
        {
            //delete all already existing sheets and related worksheetparts
            foreach (Sheet element in mergeData.Sheets)
            {
                var sheetId = element.Id;
                var sheet = target.WorkbookPart.Workbook.Descendants<Sheet>()
                           .FirstOrDefault(s => s.Id == sheetId);
				if(sheet != null)
					sheet.Remove();

                try
                {
                    var worksheetPart = (WorksheetPart)(target.WorkbookPart.GetPartById(sheetId));
                    target.WorkbookPart.DeletePart(worksheetPart);
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    continue;
                }
            }

            return target;
        }

        private void FixSheetsIds(SpreadsheetDocument target, ExcelMergeData mergeData)
        {
            uint sheetId = 1;
            foreach (var item in mergeData.Sheets)
            {
                item.Id = target.WorkbookPart.GetIdOfPart(mergeData.WorksheetPartList[item.Id]);
                //item.SheetId.Value = sheetId;
                sheetId++;
            }
        }

        private void ReplaceWorkSheetparts(SpreadsheetDocument target, ExcelMergeData mergeData)
        {
            //check make sure relationship id won`t repeat in differnet documents after adding sheets
            var worksheetPartList = mergeData.WorksheetPartList.ToList();
            foreach (KeyValuePair<string, WorksheetPart> element in worksheetPartList)
            {
                var addedWorksheetPart = target.WorkbookPart.AddPart(element.Value);
                mergeData.WorksheetPartList[element.Key] = addedWorksheetPart;
            }
        }

        private void CleanView(WorksheetPart worksheetPart)
        {
            //There can only be one sheet that has focus
            SheetViews views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }
        }

        #endregion
    }
}
