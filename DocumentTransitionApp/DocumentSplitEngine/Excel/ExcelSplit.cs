using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Excel;
using DocumentSplitEngine.Interfaces;
using OpenXmlPowerTools;
using OpenXMLTools;
using OpenXMLTools.Interfaces;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace DocumentSplitEngine
{
    public class ExcelSplit : MergeXml<Sheet>, ISplit, ISplitXml, ILocalSplit
    {
        IExcelTools ExcelTools;

        public ExcelSplit(string docName)
        {
            DocumentName = docName;
            ExcelTools = new ExcelTools();
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

            // Open a SpreadsheetDocumentDocument for editing using the filepath.
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

        public void OpenAndSearchDocument(Stream excelFile, Stream xmlFile)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            Split splitXml = (Split)serializer.Deserialize(xmlFile);
            using (SpreadsheetDocument excelDoc =
              SpreadsheetDocument.Open(excelFile, true))
            {
                Workbook workBook = excelDoc.WorkbookPart.Workbook;
                IMarkerMapper<Sheet> mapping = new MarkerExcelMapper(DocumentName, splitXml, workBook);
                DocumentElements = mapping.Run();
            }
        }

        public List<PersonFiles> SaveSplitDocument(Stream document)
        {
            List<PersonFiles> resultList = new List<PersonFiles>();

            byte[] byteArray = StreamTools.ReadFully(document);
            using (MemoryStream documentInMemoryStream = new MemoryStream(byteArray, 0, byteArray.Length, true, true))
            {
                foreach (OpenXMLDocumentPart<Sheet> element in DocumentElements)
                {
                    OpenXmlPowerToolsDocument docDividedPowerTools = new OpenXmlPowerToolsDocument(DocumentName, documentInMemoryStream);
                    using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(docDividedPowerTools))
                    {
                        SpreadsheetDocument excelDoc = streamDoc.GetSpreadsheetDocument();
                        excelDoc.WorkbookPart.Workbook.Sheets.RemoveAllChildren();
                        Sheets sheets = excelDoc.WorkbookPart.Workbook.Sheets;
                        foreach (Sheet compo in element.CompositeElements)
                            sheets.Append(compo.CloneNode(false));

                        ExcelTools.RemoveReferencedCalculationChainCell(excelDoc);
                        excelDoc.WorkbookPart.Workbook.Save();

                        var person = new PersonFiles();
                        person.Person = element.PartOwner;
                        resultList.Add(person);
                        person.Name = element.Guid.ToString();
                        person.Data = streamDoc.GetModifiedDocument().DocumentByteArray;
                    }
                }

                OpenXmlPowerToolsDocument docPowerTools = new OpenXmlPowerToolsDocument(DocumentName, documentInMemoryStream);
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                {
                    SpreadsheetDocument excelDoc = streamDoc.GetSpreadsheetDocument();

                    excelDoc.WorkbookPart.Workbook.Sheets.RemoveAllChildren();
                    excelDoc.WorkbookPart.Workbook.Save();

                    var person = new PersonFiles();
                    person.Person = "/";
                    resultList.Add(person);
                    person.Name = "template.xlsx";
                    person.Data = streamDoc.GetModifiedDocument().DocumentByteArray;
                }
            }

            var xmlPerson = new PersonFiles();
            xmlPerson.Person = "/";
            resultList.Add(xmlPerson);
            xmlPerson.Name = "mergeXmlDefinition.xml";
            xmlPerson.Data = CreateMergeXml();

            return resultList;
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
            Split splitXml;
            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            splitXml = (Split)serializer.Deserialize(xmlFile);
            var splitDocument = (SplitExcel)splitXml.Items.Where(it => it is SplitExcel && string.Equals(((SplitExcel)it).Name, DocumentName)).SingleOrDefault();
            if (splitDocument == null)
                throw new SplitNameDifferenceExcception(string.Format("This split xml describes a different document."));

            foreach (var person in splitDocument.Person)
            {
                foreach (var universalMarker in person.UniversalMarker)
                {
                    var selectedPartsIndexes = MarkerHelper<PartsSelectionTreeElement>.GetCrossedElements(universalMarker.ElementId, universalMarker.SelectionLastelementId, parts, element => element.ElementId);
                    foreach (var index in selectedPartsIndexes)
                    {
                        parts[index].OwnerName = person.Email;
                        parts[index].Selected = true;
                    }
                }
            }

            return parts;
        }
    }
}
