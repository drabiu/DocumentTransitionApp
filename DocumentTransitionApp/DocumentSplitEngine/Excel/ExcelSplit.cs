using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Excel;
using DocumentSplitEngine.Interfaces;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentSplitEngine
{
    public class ExcelSplit : MergeXml<OpenXMLDocumentPart<WorkbookPart>>, ISplit, ISplitXml, ILocalSplit
	{
        public ExcelSplit(string docName)
        {
            DocumentName = docName;
        }

        [Obsolete]
		public void OpenAndSearchDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
			//split XML Read
			var xml = File.ReadAllText(xmlSplitDefinitionFilePath);
			Split splitXml;
			using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Split));
				splitXml = (Split)serializer.Deserialize(stream);
			}

			// Open a WordprocessingDocument for editing using the filepath.
			SpreadsheetDocument wordprocessingDocument =
				SpreadsheetDocument.Open(filePath, true);

			// Assign a reference to the existing document body.
			Workbook body = wordprocessingDocument.WorkbookPart.Workbook;
			IMarkerMapper<WorkbookPart> mapping = new MarkerExcelMapper(DocumentName, splitXml, body);
			//DocumentElements = mapping.Run();

			// Close the handle explicitly.
			wordprocessingDocument.Close();
		}

        [Obsolete]
		public void SaveSplitDocument(string filePath)
		{
			throw new NotImplementedException();
		}

		public void OpenAndSearchDocument(Stream docxFile, Stream xmlFile)
		{
			throw new NotImplementedException();
		}

		public List<PersonFiles> SaveSplitDocument(Stream document)
		{
			throw new NotImplementedException();
		}

        public byte[] CreateSplitXml(IList<PartsSelectionTreeElement> parts)
        {
            throw new NotImplementedException();
        }

        public List<PartsSelectionTreeElement> PartsFromSplitXml(Stream xmlFile, List<PartsSelectionTreeElement> parts)
        {
            throw new NotImplementedException();
        }
    }
}
