using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentMergeEngine.Interfaces;
using OpenXmlPowerTools;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocumentMergeEngine
{
    public class ExcelMerge : DocumentMerge, IMerge
    {
        public byte[] Run(List<PersonFiles> files)
        {
            var mergeXml = GetMergeXml(files);

            Sheets sheets = new Sheets();
            MergeDocument documentXml = mergeXml.Items.First();
            foreach (MergeDocumentPart part in documentXml.Part)
            {
                byte[] byteArray = files.Where(p => p.Person == part.Name && p.Name == part.Id).Select(d => d.Data).FirstOrDefault();
                using (MemoryStream docInMemoryStream = new MemoryStream(byteArray, 0, byteArray.Length, true, true))
                {
                    OpenXmlPowerToolsDocument docPowerTools = new OpenXmlPowerToolsDocument(string.Empty, docInMemoryStream);
                    using (OpenXmlMemoryStreamDocument streamEmptyDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                    {
                        SpreadsheetDocument spreadSheetDocument = streamEmptyDoc.GetSpreadsheetDocument();

                        // Assign a reference to the existing document body.
                        foreach (Sheet element in spreadSheetDocument.WorkbookPart.Workbook.Sheets)
                        {
                            sheets.Append(element.CloneNode(true));
                        }

                        // Close the handle explicitly.
                        spreadSheetDocument.Close();
                    }
                }
            }

            byte[] template = files.Where(p => p.Person == "/" && p.Name == "template.xlsx").Select(d => d.Data).FirstOrDefault();
            using (MemoryStream docInMemoryStream = new MemoryStream(template, 0, template.Length, true, true))
            {
                OpenXmlPowerToolsDocument docPowerTools = new OpenXmlPowerToolsDocument(string.Empty, docInMemoryStream);
                using (OpenXmlMemoryStreamDocument streamEmptyDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                {
                    SpreadsheetDocument excelDoc = streamEmptyDoc.GetSpreadsheetDocument();
                    excelDoc.WorkbookPart.Workbook.Sheets = sheets;
                    excelDoc.WorkbookPart.Workbook.Save();

                    return streamEmptyDoc.GetModifiedDocument().DocumentByteArray;
                }
            }
        }
    }
}
