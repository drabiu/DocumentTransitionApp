using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentEditPartsEngine.Interfaces;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentEditPartsEngine
{
    public static class ExcelDocumentPartAttributes
    {
        public const int MaxNameLength = 36;

        public static string GetSlideIdFormatter(int id)
        {
            return string.Format("slId{0}", id);
        }

        public static bool IsSupportedType(OpenXmlElement part)
        {
            bool isSupported = false;
            isSupported = part is Sheet;
            //|| element is Wordproc.Picture
            //|| element is Wordproc.Drawing
            //|| element is Wordproc.Table;

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
                    result.Add(new PartsSelectionTreeElement(id.ToString(), ExcelDocumentPartAttributes.GetSlideIdFormatter(id), sheetName, 0));
                }
            }

            return result;
        }
    }
}
