using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentMergeEngine.Interfaces;
using OpenXmlPowerTools;
using OpenXMLTools;
using OpenXMLTools.Interfaces;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocumentMergeEngine
{
    public class ExcelMerge : DocumentMerge, IMerge
    {
        IExcelTools ExcelTools;

        public byte[] Run(List<PersonFiles> files)
        {
            var mergeXml = GetMergeXml(files);

            byte[] template = files.Where(p => p.Person == "/" && p.Name == "template.xlsx").Select(d => d.Data).FirstOrDefault();
            using (MemoryStream docInMemoryStream = new MemoryStream(template, 0, template.Length, true, true))
            {
                OpenXmlPowerToolsDocument docPowerTools = new OpenXmlPowerToolsDocument(string.Empty, docInMemoryStream);
                using (OpenXmlMemoryStreamDocument streamEmptyDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                {
                    SpreadsheetDocument excelTemplateDoc = streamEmptyDoc.GetSpreadsheetDocument();
                    ExcelTools = new ExcelTools(excelTemplateDoc);
                    Sheets sheets = new Sheets();
                    List<SharedStringItem> sharedStringItems = new List<SharedStringItem>();

                    MergeDocument documentXml = mergeXml.Items.First();
                    foreach (MergeDocumentPart part in documentXml.Part)
                    {
                        byte[] byteArray = files.Where(p => p.Person == part.Name && p.Name == part.Id).Select(d => d.Data).FirstOrDefault();
                        using (MemoryStream docPartInMemoryStream = new MemoryStream(byteArray, 0, byteArray.Length, true, true))
                        {
                            OpenXmlPowerToolsDocument docPartPowerTools = new OpenXmlPowerToolsDocument(string.Empty, docPartInMemoryStream);
                            using (OpenXmlMemoryStreamDocument streamPartDoc = new OpenXmlMemoryStreamDocument(docPartPowerTools))
                            {                     
                                SpreadsheetDocument spreadSheetDocument = streamPartDoc.GetSpreadsheetDocument();

                                sharedStringItems.AddRange(ExcelTools.GetAddedSharedStringItems(spreadSheetDocument));
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


                    excelTemplateDoc.WorkbookPart.Workbook.Sheets = sheets;
                    excelTemplateDoc.WorkbookPart.SharedStringTablePart.SharedStringTable.Append(sharedStringItems);
                    excelTemplateDoc.WorkbookPart.Workbook.Save();

                    return streamEmptyDoc.GetModifiedDocument().DocumentByteArray;
                }
            }
        }
    }
}
