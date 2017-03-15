using DocumentEditPartsEngine.Helpers;
using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngine
{
    public static class ExcelDocumentPartAttributes
    {
        public const int MaxNameLength = 36;

        public static string GetSheetIdFormatter(int id)
        {
            return string.Format("shId{0}", id);
        }

        public static bool IsSupportedType(OpenXmlElement element)
        {
            bool isSupported = false;
            isSupported = element is Sheet
            || element is Row
            //|| element is Column
            || element is Cell;

            return isSupported;
        }
    }

    public class ExcelDocumentParts : IExcelParts
    {
        public List<PartsSelectionTreeElement> Get(Stream file, Predicate<OpenXmlElement> supportedTypes)
        {
            List<PartsSelectionTreeElement> workBookElements = new List<PartsSelectionTreeElement>();
            using (SpreadsheetDocument excDoc =
                SpreadsheetDocument.Open(file, true))
            {
                Workbook workBook = excDoc.WorkbookPart.Workbook;
                var idIndex = 1;
                for (int cellIndex = 0; cellIndex < workBook.Sheets.ChildElements.Count; cellIndex++)
                {
                    Sheet sheet = workBook.Sheets.ChildElements[cellIndex] as Sheet;
                    workBookElements.AddRange(CreatePartsSelectionTreeElements(sheet, cellIndex + 1, supportedTypes));
                    idIndex++;

                    var worksheetPart = (WorksheetPart)(excDoc.WorkbookPart.GetPartById(sheet.Id));
                    //foreach (var row in worksheetPart.Worksheet.Descendants<Column>())
                    //{
                    //    idIndex++;
                    //}

                    foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
                    {
                        workBookElements.AddRange(CreatePartsSelectionTreeElements(row, idIndex, supportedTypes));
                        idIndex++;
                        foreach (var cell in row.Elements<Cell>())
                        {
                            workBookElements.AddRange(CreatePartsSelectionTreeElements(cell, idIndex, supportedTypes));
                            idIndex++;
                        }
                    }
                }
            }

            return workBookElements;
        }

        public List<PartsSelectionTreeElement> GetSheets(Stream file)
        {
            return Get(file, el => ExcelDocumentPartAttributes.IsSupportedType(el));
        }

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, int id, Predicate<OpenXmlElement> isSupportedType)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (isSupportedType(element))
            {
                if (element is Sheet)
                {
                    string sheetName = string.Format("{0}", (element as Sheet).Name);
                    string elementId = ExcelDocumentPartAttributes.GetSheetIdFormatter(id);
                    result.Add(new PartsSelectionTreeElement(elementId, elementId, sheetName, 0, ElementType.Sheet));
                }
                else if (element is Row)
                {
                    Row row = element as Row;
                    string rowName = string.Format("Row index: {0}", row.RowIndex);
                    result.Add(new PartsSelectionTreeElement(id.ToString(), row.RowIndex, rowName, 1, ElementType.Row));
                }
                else if (element is Column)
                {

                }
                else if (element is Cell)
                {
                    Cell cell = element as Cell;
                    string cellName = string.Format("Cell name: {0}", cell.CellReference.Value);
                    result.Add(new PartsSelectionTreeElement(id.ToString(), cell.CellReference.Value, cellName, 2, ElementType.Cell));
                }
            }

            return result;
        }
    }
}
