using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Linq;
using DocumentEditPartsEngine;

namespace SplitDescriptionObjects
{
    public interface IExcelMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedSheetElements(string id, string id2);
    }

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
                result = ExcelDocumentPartAttributes.GetSlideIdFormatter(index + 1);
            }

            return result;
        }
    }

    public interface IUniversalExcelMarker : IExcelMarker
    {
    }

    public class UniversalExcelMarker : ExcelMarker, IUniversalExcelMarker
    {
        public UniversalExcelMarker(Workbook body) :
            base(body)
        {
        }
    }
}
