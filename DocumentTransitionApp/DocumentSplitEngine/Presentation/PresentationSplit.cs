using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using DocumentSplitEngine.Presentation;
using OpenXmlPowerTools;
using OpenXMLTools;
using OpenXMLTools.Interfaces;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace DocumentSplitEngine
{
    public class PresentationSplit : MergeXml<SlideId>, ISplit, ISplitXml, ILocalSplit
	{
        IPresentationTools PresentationTools;

        public PresentationSplit(string docName)
        {
            DocumentName = docName;
            PresentationTools = new PresentationTools();
        }

        [Obsolete]
        public void OpenAndSearchDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
            throw new NotImplementedException();
        }

        [Obsolete]
        public void SaveSplitDocument(string filePath)
		{
			throw new NotImplementedException();
		}

        public void OpenAndSearchDocument(Stream docFile, Stream xmlFile)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            Split splitXml = (Split)serializer.Deserialize(xmlFile);
            using (PresentationDocument preDoc =
              PresentationDocument.Open(docFile, true))
            {
                PresentationPart body = preDoc.PresentationPart;
                IMarkerMapper<SlideId> mapping = new MarkerPresentationMapper(DocumentName, splitXml, body);
                DocumentElements = mapping.Run();
            }
        }

		public List<PersonFiles> SaveSplitDocument(Stream document)
		{
            List<PersonFiles> resultList = new List<PersonFiles>();

            byte[] byteArray = StreamTools.ReadFully(document);
            using (MemoryStream documentInMemoryStream = new MemoryStream(byteArray, 0, byteArray.Length, true, true))
            {
                OpenXmlPowerToolsDocument docPowerTools = new OpenXmlPowerToolsDocument(DocumentName, documentInMemoryStream);
                using (OpenXmlMemoryStreamDocument streamTemplateDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                {
                    PresentationDocument templatePresentation = streamTemplateDoc.GetPresentationDocument();
                    foreach (OpenXMLDocumentPart<SlideId> element in DocumentElements)
                    {
                        OpenXmlPowerToolsDocument emptyDocPowerTools = new OpenXmlPowerToolsDocument(DocumentName, documentInMemoryStream);
                        using (OpenXmlMemoryStreamDocument streamDividedDoc = new OpenXmlMemoryStreamDocument(emptyDocPowerTools))
                        {
                            var relationshipIds = element.CompositeElements.Select(c => c.RelationshipId.Value).ToList();
                            PresentationDocument dividedPresentation = streamDividedDoc.GetPresentationDocument();
                            PresentationTools.InsertSlidesFromTemplate(PresentationTools.RemoveAllSlides(dividedPresentation), templatePresentation, relationshipIds);

                            var person = new PersonFiles();
                            person.Person = element.PartOwner;
                            resultList.Add(person);
                            person.Name = element.Guid.ToString();
                            person.Data = streamDividedDoc.GetModifiedDocument().DocumentByteArray;
                        }
                    }
                }

                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(docPowerTools))
                {
                    PresentationDocument preDoc = streamDoc.GetPresentationDocument();
                    PresentationTools.RemoveAllSlides(preDoc);

                    var person = new PersonFiles();
                    person.Person = "/";
                    resultList.Add(person);
                    person.Name = "template.pptx";
                    person.Data = streamDoc.GetModifiedDocument().DocumentByteArray;
                }

                var xmlPerson = new PersonFiles();
                xmlPerson.Person = "/";
                resultList.Add(xmlPerson);
                xmlPerson.Name = "mergeXmlDefinition.xml";
                xmlPerson.Data = CreateMergeXml();
            }

            return resultList;
        }

        public byte[] CreateSplitXml(IList<PartsSelectionTreeElement> parts)
        {
            var nameList = parts.Select(p => p.OwnerName).Where(n => !string.IsNullOrEmpty(n)).Distinct().ToList();
            var indexer = new NameIndexer(nameList);

            Split splitXml = new Split();
            splitXml.Items = new SplitPresentation[1];
            splitXml.Items[0] = new SplitPresentation();
            (splitXml.Items[0] as SplitPresentation).Name = DocumentName;
            var splitDocument = (splitXml.Items[0] as SplitPresentation);
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
            var splitDocument = (SplitPresentation)splitXml.Items.Where(it => it is SplitPresentation && string.Equals(((SplitPresentation)it).Name, DocumentName)).SingleOrDefault();
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

        #region Private methods
        #endregion
    }
}
