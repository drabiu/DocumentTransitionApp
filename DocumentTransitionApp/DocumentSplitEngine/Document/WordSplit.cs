using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Reflection;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using SplitDescriptionObjects;
using DocumentEditPartsEngine;
using DocumentSplitEngine.Data_Structures;
using OpenXMLTools;
using DocumentSplitEngine.Document;
using DocumentSplitEngine.Interfaces;
using OpenXmlPowerTools;

namespace DocumentSplitEngine
{
	public class WordSplit : MergeXml<OpenXmlElement>, ISplit, ISplitXml, ILocalSplit
    {       
		public WordSplit(string docName)
		{
			DocumentName = docName;
		}
      
        public void OpenAndSearchDocument(Stream docxFile, Stream xmlFile)
		{
			XmlSerializer serializer = new XmlSerializer(typeof(Split));
			Split splitXml = (Split)serializer.Deserialize(xmlFile);

            byte[] byteArray = StreamTools.ReadFully(docxFile);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(mem, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    IMarkerMapper<OpenXmlElement> mapping = new MarkerWordMapper(DocumentName, splitXml, body);
                    DocumentElements = mapping.Run();
                }
            }
		}

        [Obsolete]
		public void OpenAndSearchDocument(string docxFilePath, string xmlFilePath)
		{
			//split XML Read
			var xml = File.ReadAllText(xmlFilePath);
			Split splitXml;
			using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Split));
				splitXml = (Split)serializer.Deserialize(stream);
			}

			// Open a WordprocessingDocument for editing using the filepath.
			WordprocessingDocument wordprocessingDocument =
				WordprocessingDocument.Open(docxFilePath, true);

			// Assign a reference to the existing document body.
			Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
			MarkerWordMapper mapping = new MarkerWordMapper(DocumentName, splitXml, body);
			DocumentElements = mapping.Run();

			// Close the handle explicitly.
			wordprocessingDocument.Close();
		}

        [Obsolete]
		public void SaveSplitDocument(string docxFilePath)
		{
			DirectoryInfo initDi;
			string appPath = Path.GetDirectoryName(Assembly.GetAssembly(typeof(WordSplit)).Location);
			if (!Directory.Exists(appPath + @"\Files"))
				initDi = Directory.CreateDirectory(appPath + @"\Files");

			byte[] byteArray = File.ReadAllBytes(docxFilePath);
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					foreach (OpenXMLDocumentPart<OpenXmlElement> element in DocumentElements)
					{						
						wordDoc.MainDocumentPart.Document.Body = new Body();
						Body body = wordDoc.MainDocumentPart.Document.Body;
						foreach (OpenXmlElement compo in element.CompositeElements)
							body.Append(compo.CloneNode(true));

						wordDoc.MainDocumentPart.Document.Save();

						string directoryPath = appPath + @"\Files" + @"\" + element.PartOwner;
						DirectoryInfo currentDi;
						if (!Directory.Exists(directoryPath))
						{
							currentDi = Directory.CreateDirectory(directoryPath);
						}

						using (FileStream fileStream = new FileStream(directoryPath + @"\" + element.Guid.ToString() + ".docx",
                            FileMode.CreateNew))
						{
							mem.WriteTo(fileStream);
						}
					}
				}
				// At this point, the memory stream contains the modified document.
				// We could write it back to a SharePoint document library or serve
				// it from a web server.			
			}

			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					wordDoc.MainDocumentPart.Document.Body = new Body();
					wordDoc.MainDocumentPart.Document.Save();

					using (FileStream fileStream = new FileStream(appPath + @"\Files" + @"\template" + ".docx",
                        FileMode.CreateNew))
					{
						mem.WriteTo(fileStream);
					}
				}
				// At this point, the memory stream contains the modified document.
				// We could write it back to a SharePoint document library or serve
				// it from a web server.			
			}

			CreateMergeXml(appPath + @"\Files" + @"\");
		}

		public List<PersonFiles> SaveSplitDocument(Stream document)
		{
			List<PersonFiles> resultList = new List<PersonFiles>();

			byte[] byteArray = StreamTools.ReadFully(document);
            using (MemoryStream documentInMemoryStream = new MemoryStream(byteArray, 0, byteArray.Length, true, true))
            {
                foreach (OpenXMLDocumentPart<OpenXmlElement> element in DocumentElements)
                {
                    OpenXmlPowerToolsDocument docDividedPowerTools = new OpenXmlPowerToolsDocument(DocumentName, documentInMemoryStream);
                    using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(docDividedPowerTools))
                    {
                        WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument();
                        wordDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                        Body body = wordDoc.MainDocumentPart.Document.Body;
                        foreach (OpenXmlElement compo in element.CompositeElements)
                            body.Append(compo.CloneNode(true));

                        wordDoc.MainDocumentPart.Document.Save();

                        var person = new PersonFiles();
                        person.Person = element.PartOwner;
                        resultList.Add(person);
                        person.Name = element.Guid.ToString();
                        person.Data = streamDoc.GetModifiedDocument().DocumentByteArray;

                    }
			    }
                // At this point, the memory stream contains the modified document.
                // We could write it back to a SharePoint document library or serve
                // it from a web server.	

                OpenXmlPowerToolsDocument docPowerTools = new OpenXmlPowerToolsDocument(DocumentName, documentInMemoryStream);
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                {
                    WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument();

                    wordDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    wordDoc.MainDocumentPart.Document.Save();

                    var person = new PersonFiles();
                    person.Person = "/";
                    resultList.Add(person);
                    person.Name = "template.docx";
                    person.Data = streamDoc.GetModifiedDocument().DocumentByteArray;
                }
            }
			// At this point, the memory stream contains the modified document.
			// We could write it back to a SharePoint document library or serve
			// it from a web server.			

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
            splitXml.Items = new SplitDocument[1];
            splitXml.Items[0] = new SplitDocument();
            (splitXml.Items[0] as SplitDocument).Name = DocumentName;
            var splitDocument = (splitXml.Items[0] as SplitDocument);
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
            var splitDocument = (SplitDocument)splitXml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, DocumentName)).SingleOrDefault();
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
