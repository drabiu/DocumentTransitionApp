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

        public static string GetSlideIdFormatter(int id)
        {
            return string.Format("slId{0}", id);
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
                foreach (Sheet sheet in workBook.Sheets)
                {
                    workBookElements.AddRange(CreatePartsSelectionTreeElements(sheet, idIndex, supportedTypes));
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
                    result.Add(new PartsSelectionTreeElement(id.ToString(), ExcelDocumentPartAttributes.GetSlideIdFormatter(id), sheetName, 0, ElementType.Sheet));
                }
                else if (element is Row)
                {
                    string rowName = string.Format("Row index: {0}", (element as Row).RowIndex);
                    result.Add(new PartsSelectionTreeElement(id.ToString(), (element as Row).RowIndex, rowName, 1, ElementType.Row));
                }
                else if (element is Column)
                {

                }
                else if (element is Cell)
                {
                    string cellName = string.Format("Cell name: {0}", (element as Cell).CellReference.Value);
                    result.Add(new PartsSelectionTreeElement(id.ToString(), (element as Cell).CellReference.Value, cellName, 2, ElementType.Cell));
                }
            }

            return result;
        }
    }
}
