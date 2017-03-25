using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SplitDescriptionObjects.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects
{
    public abstract class ExcelMarker : IExcelMarker
    {
        Workbook DocumentBody;
        List<Sheet> ElementsList;

        public ExcelMarker(Workbook body)
        {
            DocumentBody = body;
            ElementsList = DocumentBody.WorkbookPart.Workbook.Sheets.Elements<Sheet>().ToList();
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        public IList<int> GetCrossedSheetElements(string id, string id2)
        {
            var indexes = MarkerHelper<Sheet>.GetCrossedElements(id, id2, ElementsList, element => GetSheetId(element));

            return indexes;
        }

        private string GetSheetId(OpenXmlElement element)
        {
            string result = string.Empty;
            if (element is Sheet)
            {
                Sheet sheet = (element as Sheet);
                int index = ElementsList.FindIndex(el => el.Equals(element));
                result = ExcelDocumentPartAttributes.GetSheetIdFormatter(index + 1);
            }

            return result;
        }
    }

    public class SheetExcelMarker : ExcelMarker, ISheetExcelMarker
    {
        public SheetExcelMarker(Workbook body) :
            base(body)
        {
        }
    }
}
