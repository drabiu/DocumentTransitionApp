using DocumentFormat.OpenXml.Spreadsheet;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using SplitDescriptionObjects;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace DocumentSplitEngine.Excel
{
    public class MarkerExcelMapper : MarkerDocumentMapper, IMarkerMapper<Sheet>
    {
        SplitExcel SplitExcelObj { get; set; }
        Workbook WorkBook;
        ISheetExcelMarker UniversalExcMarker;

        public MarkerExcelMapper(string documentName, Split xml, Workbook workBook)
        {
            Xml = xml;
            SplitExcelObj = (SplitExcel)Xml.Items.Where(it => it is SplitExcel && string.Equals(((SplitExcel)it).Name, documentName)).SingleOrDefault();
            WorkBook = workBook;
            UniversalExcMarker = new SheetExcelMarker(workBook);
            SubdividedParagraphs = new string[workBook.Sheets.Count()];
        }

        public IList<OpenXMLDocumentPart<Sheet>> Run()
        {
            IList<OpenXMLDocumentPart<Sheet>> documentElements = new List<OpenXMLDocumentPart<Sheet>>();
            if (SplitExcelObj != null)
            {
                foreach (Person person in SplitExcelObj.Person)
                {
                    if (person.UniversalMarker != null)
                    {
                        foreach (PersonUniversalMarker marker in person.UniversalMarker)
                        {
                            IList<int> result = UniversalExcMarker.GetCrossedSheetElements(marker.ElementId, marker.SelectionLastelementId);
                            foreach (int index in result)
                            {
                                if (string.IsNullOrEmpty(SubdividedParagraphs[index]))
                                {
                                    SubdividedParagraphs[index] = person.Email;
                                }
                                else
                                    throw new ElementToPersonPairException();
                            }
                        }
                    }
                }

                string email = string.Empty;
                OpenXMLDocumentPart<Sheet> part = new OpenXMLDocumentPart<Sheet>();
                var sheetPartsList = WorkBook.Sheets.Elements<Sheet>().ToList();
                for (int index = 0; index < sheetPartsList.Count; index++)
                {
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<Sheet>();
                        part.CompositeElements.Add(sheetPartsList[index] as Sheet);
                        email = SubdividedParagraphs[index];
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                        part.CompositeElements.Add(sheetPartsList[index] as Sheet);
                }
            }

            return documentElements;
        }
    }
}
