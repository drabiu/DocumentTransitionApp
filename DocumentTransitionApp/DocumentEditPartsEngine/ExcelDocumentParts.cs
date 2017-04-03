using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLTools.Helpers;
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
            isSupported = element is Sheet;
            //|| element is Row
            //|| element is Column
            //|| element is Cell;

            return isSupported;
        }
    }

    public class ExcelDocumentParts : IDocumentParts
    {
        DocumentParts _documentParts;

        public ExcelDocumentParts()
        {
            _documentParts = new DocumentParts(this);
        }

        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            return Get(file, el => ExcelDocumentPartAttributes.IsSupportedType(el));
        }

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
                    workBookElements.AddRange(_documentParts.CreatePartsSelectionTreeElements(sheet, null, cellIndex + 1, supportedTypes, 0, true));
                    idIndex++;

                    //var worksheetPart = (WorksheetPart)(excDoc.WorkbookPart.GetPartById(sheet.Id));
                    //foreach (var row in worksheetPart.Worksheet.Descendants<Column>())
                    //{
                    //    idIndex++;
                    //}

                    //foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
                    //{
                    //    workBookElements.AddRange(_documentParts.CreatePartsSelectionTreeElements(row, null, idIndex, supportedTypes, 1));
                    //    idIndex++;
                    //    foreach (var cell in row.Elements<Cell>())
                    //    {
                    //        workBookElements.AddRange(_documentParts.CreatePartsSelectionTreeElements(cell, null, idIndex, supportedTypes, 2));
                    //        idIndex++;
                    //    }
                    //}
                }
            }

            return workBookElements;
        }

        public PartsSelectionTreeElement GetParagraphSelectionTreeElement(OpenXmlElement element, PartsSelectionTreeElement parent, ref int id, Predicate<OpenXmlElement> supportedType, int indent, bool visible)
        {
            PartsSelectionTreeElement elementToAdd = null;
            if (supportedType(element))
            {
                if (element is Sheet)
                {
                    string sheetName = string.Format("{0}", (element as Sheet).Name);
                    string elementId = ExcelDocumentPartAttributes.GetSheetIdFormatter(id);
                    elementToAdd = new PartsSelectionTreeElement(elementId, elementId, sheetName, indent, ElementType.Sheet);
                }
                else if (element is Row)
                {
                    Row row = element as Row;
                    string rowName = string.Format("Row index: {0}", row.RowIndex);
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), row.RowIndex, rowName, indent, ElementType.Row);
                }
                else if (element is Column)
                {

                }
                else if (element is Cell)
                {
                    Cell cell = element as Cell;
                    string cellName = string.Format("Cell name: {0}", cell.CellReference.Value);
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), cell.CellReference.Value, cellName, indent, ElementType.Cell);
                }
            }

            return elementToAdd;
        }
    }
}
