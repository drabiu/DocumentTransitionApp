using DocumentFormat.OpenXml.Packaging;
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

        public void AppendToSharedStringTablePart()
        {
            //need to add string to this table in order for them to be merged later
            //excelDoc.WorkbookPart.SharedStringTablePart
        }

        public void RemoveReferencesFromCalculationChainPart()
        {
            //when spliting need to remove unused references to removed sheets
            //excelDoc.WorkbookPart.CalculationChainPart 
        }

        #endregion
    }
}
