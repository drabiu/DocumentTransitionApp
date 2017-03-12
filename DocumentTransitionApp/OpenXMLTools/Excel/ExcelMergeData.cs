using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace OpenXMLTools
{
    public class ExcelMergeData
    {
        public List<Sheet> Sheets;
        public Dictionary<string, WorksheetPart> WorksheetPartList;

        public ExcelMergeData()
        {
            InitFields();
        }

        public void AppendDocumentData(SpreadsheetDocument document)
        {
            foreach (Sheet element in document.WorkbookPart.Workbook.Sheets)
            {
                Sheets.Add(element.CloneNode(true) as Sheet);
            }

            foreach (WorksheetPart element in document.WorkbookPart.WorksheetParts)
            {
                string elemntId = document.WorkbookPart.GetIdOfPart(element);
                //this simple validation makes sure that we don`t add any workbookpart without relation to sheet
                if (document.WorkbookPart.Workbook.Sheets.Any(s => (s as Sheet).Id == elemntId))
                    WorksheetPartList.Add(elemntId, element);
            }
        }

        public void SetDocumentData(SpreadsheetDocument document)
        {
            InitFields();
            AppendDocumentData(document);
        }

        private void InitFields()
        {
            Sheets = new List<Sheet>();
            WorksheetPartList = new Dictionary<string, WorksheetPart>();
        }
    }
}
