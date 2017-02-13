using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Excelproc = DocumentFormat.OpenXml.Spreadsheet;

using SplitDescriptionObjects;
using DocumentEditPartsEngine;
using DocumentSplitEngine.Data_Structures;

namespace DocumentSplitEngine
{
    public interface IExcelMarkerMapper
    {
        IList<OpenXMLDocumentPart<WorkbookPart>> Run();
    }

	public class MarkerExcelMapper : MarkerMapper, IExcelMarkerMapper
    {
		SplitExcel SplitExcelObj { get; set; }
		Excelproc.Workbook WorkBook;

		public MarkerExcelMapper(string documentName, Split xml, Excelproc.Workbook workBook)
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

	public class ExcelSplit : MergeXml<OpenXMLDocumentPart<WorkbookPart>>, ISplit, ILocalSplit
	{
        [Obsolete]
		public void OpenAndSearchDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
			//split XML Read
			var xml = System.IO.File.ReadAllText(xmlSplitDefinitionFilePath);
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
			Excelproc.Workbook body = wordprocessingDocument.WorkbookPart.Workbook;
			IExcelMarkerMapper mapping = new MarkerExcelMapper(DocumentName, splitXml, body);
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

		List<PersonFiles> ISplit.SaveSplitDocument(Stream document)
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
