using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLTools.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class ExcelTools : IExcelTools
    {
        #region Public methods

        SpreadsheetDocument _excelDocument;

        public ExcelTools(SpreadsheetDocument document)
        {
            _excelDocument = document;
        }

        public IEnumerable<SharedStringItem> GetAddedSharedStringItems(SpreadsheetDocument target)
        {
            var items = target.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Select(el => el.CloneNode(true) as SharedStringItem);

            return items;
        }

        public void RemoveReferencesFromCalculationChainPart()
        {
            //when spliting need to remove unused references to removed sheets
            //excelDoc.WorkbookPart.CalculationChainPart 
        }

        #endregion
    }
}
