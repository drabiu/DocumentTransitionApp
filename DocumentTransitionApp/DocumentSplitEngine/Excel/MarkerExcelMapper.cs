using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace DocumentSplitEngine.Excel
{
    public class MarkerExcelMapper : MarkerDocumentMapper, IMarkerMapper<Sheet>
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

        public IList<OpenXMLDocumentPart<Sheet>> Run()
        {
            IList<OpenXMLDocumentPart<Sheet>> documentElements = new List<OpenXMLDocumentPart<Sheet>>();
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
                OpenXMLDocumentPart<Sheet> part = new OpenXMLDocumentPart<Sheet>();
                for (int index = 0; index < WorkBook.ChildElements.Count; index++)
                {
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<Sheet>();
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
