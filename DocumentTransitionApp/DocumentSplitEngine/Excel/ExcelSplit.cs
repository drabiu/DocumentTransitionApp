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
using System.Linq;

namespace DocumentSplitEngine
{
    public class ExcelSplit : MergeXml<OpenXMLDocumentPart<Sheet>>, ISplit, ISplitXml, ILocalSplit
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
			IMarkerMapper<Sheet> mapping = new MarkerExcelMapper(DocumentName, splitXml, body);
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
            var nameList = parts.Select(p => p.OwnerName).Where(n => !string.IsNullOrEmpty(n)).Distinct().ToList();
            var indexer = new NameIndexer(nameList);

            Split splitXml = new Split();
            splitXml.Items = new SplitExcel[1];
            splitXml.Items[0] = new SplitExcel();
            (splitXml.Items[0] as SplitExcel).Name = DocumentName;
            var splitDocument = (splitXml.Items[0] as SplitExcel);
            splitDocument.Person = new Person[nameList.Count];
            foreach (var name in nameList)
            {
                var person = new Person();
                person.Email = name;
                person.UniversalMarker = new PersonUniversalMarker[parts.Where(p => p.OwnerName == name).Count()];
                splitDocument.Person[nameList.IndexOf(name)] = person;

            }

            foreach (var part in parts.Where(p => !string.IsNullOrEmpty(p.OwnerName)))
            {
                var person = splitDocument.Person[nameList.IndexOf(part.OwnerName)];
                var universalMarker = new PersonUniversalMarker();
                universalMarker.ElementId = part.ElementId;
                universalMarker.SelectionLastelementId = part.ElementId;
                person.UniversalMarker[indexer.GetNextIndex(part.OwnerName)] = universalMarker;
            }

            using (MemoryStream splitStream = new MemoryStream())
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Split));
                serializer.Serialize(splitStream, splitXml);

                return splitStream.ToArray();
            }
        }

        public List<PartsSelectionTreeElement> SelectPartsFromSplitXml(Stream xmlFile, List<PartsSelectionTreeElement> parts)
        {
            throw new NotImplementedException();
        }
    }
}
