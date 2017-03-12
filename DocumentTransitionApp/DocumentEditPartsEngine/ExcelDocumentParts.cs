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
            || element is Column;

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
                foreach (var sheet in workBook.Sheets)
                {
                    workBookElements.AddRange(CreatePartsSelectionTreeElements(sheet, idIndex, supportedTypes));
                    idIndex++;
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
                    string sheetName = string.Format("[Sht]: {0}", (element as Sheet).Name);
                    result.Add(new PartsSelectionTreeElement(id.ToString(), ExcelDocumentPartAttributes.GetSlideIdFormatter(id), sheetName, 0, new ExcelElementType.SheetElementSubType()));
                }
            }

            return result;
        }
    }
}
