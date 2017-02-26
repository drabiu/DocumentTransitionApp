using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace DocumentSplitEngine.Excel
{
    public class MarkerExcelMapper : MarkerMapper, IMarkerMapper<WorkbookPart>
    {
        SplitExcel SplitExcelObj { get; set; }
        Workbook WorkBook;

        public MarkerExcelMapper(string documentName, Split xml, Workbook workBook)
        {
            Xml = xml;
            SplitExcelObj = (SplitExcel)Xml.Items.Where(it => it is SplitExcel && string.Equals(((SplitExcel)it).Name, documentName)).SingleOrDefault();
            WorkBook = workBook;
            SubdividedParagraphs = new string[workBook.ChildElements.Count];
        }

        public IList<OpenXMLDocumentPart<WorkbookPart>> Run()
        {
            IList<OpenXMLDocumentPart<WorkbookPart>> documentElements = new List<OpenXMLDocumentPart<WorkbookPart>>();
            if (SplitExcelObj != null)
            {
                foreach (Person person in SplitExcelObj.Person)
                {
                    if (person.SheetMarker != null)
                    {
                        foreach (PersonSheetMarker marker in person.SheetMarker)
                        {
                            //int result = GetSheetMarker().FindElement(marker.ElementId);
                            //if (string.IsNullOrEmpty(SubdividedParagraphs[result]))
                            //{
                            //	SubdividedParagraphs[result] = person.Email;
                            //}
                            //else
                            //	throw new ElementToPersonPairException();
                        }
                    }
                }

                string email = string.Empty;
                OpenXMLDocumentPart<WorkbookPart> part = new OpenXMLDocumentPart<WorkbookPart>();
                for (int index = 0; index < WorkBook.ChildElements.Count; index++)
                {
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<WorkbookPart>();
                        //part.CompositeElements.Add(WorkBook.ChildElements[index]);
                        email = SubdividedParagraphs[index];
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                    { }
                    //part.CompositeElements.Add(WorkBook.ChildElements[index]);
                }
            }

            return documentElements;
        }
    }
}
