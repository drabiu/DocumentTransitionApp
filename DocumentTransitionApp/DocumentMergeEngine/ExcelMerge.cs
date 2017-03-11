using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentMergeEngine.Interfaces;
using OpenXmlPowerTools;
using OpenXMLTools;
using OpenXMLTools.Interfaces;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocumentMergeEngine
{
    public class ExcelMerge : DocumentMerge, IMerge
    {
        IExcelTools ExcelTools;

        public ExcelMerge()
        {
            ExcelTools = new ExcelTools();
        }

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
                    excelTemplateDoc.WorkbookPart.Workbook.Sheets = new Sheets();

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
                                ExcelTools.MergeWorkSheets(excelTemplateDoc, spreadSheetDocument);

                                // Close the handle explicitly.
                                spreadSheetDocument.Close();
                            }
                        }
                    }
                                                
                    return streamEmptyDoc.GetModifiedDocument().DocumentByteArray;
                }
            }
        }      
    }
}
