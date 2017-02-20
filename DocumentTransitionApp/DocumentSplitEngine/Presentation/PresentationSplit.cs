using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentSplitEngine.Data_Structures;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Presentation;
using OpenXMLTools;
using OpenXMLTools.Interfaces;

namespace DocumentSplitEngine
{
    public interface IPresentationMarkerMapper
    {
        IList<OpenXMLDocumentPart<SlideId>> Run();
    }

    public class MarkerPresentationMapper : MarkerMapper, IPresentationMarkerMapper
	{
        SplitDocument SplitPresentationObj { get; set; }
        PresentationPart Presentation;

		public MarkerPresentationMapper(string documentName, Split xml, PresentationPart presentation)
		{
			Xml = xml;
            SplitPresentationObj = (SplitDocument)Xml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, documentName)).SingleOrDefault();
            //SplitPresentationObj = (SplitPresentation)Xml.Items.Where(it => it is SplitPresentation && string.Equals(((SplitPresentation)it).Name, documentName)).SingleOrDefault();
            Presentation = presentation;
			SubdividedParagraphs = new string[presentation.SlideParts.Count()];
		}

        public IUniversalPresentationMarker GetUniversalDocumentMarker()
        {
            return new UniversalPresentationMarker(Presentation);
        }

        /// <summary>
        /// Finds parts of documents selected by the marker and returns as a list of persons each containing list of document elements
        /// </summary>
        /// <returns></returns>
        public IList<OpenXMLDocumentPart<SlideId>> Run()
		{
            IList<OpenXMLDocumentPart<SlideId>> documentElements = new List<OpenXMLDocumentPart<SlideId>>();
            if (SplitPresentationObj != null)
            {
                foreach (Person person in SplitPresentationObj.Person)
                {
                    if (person.UniversalMarker != null)
                    {
                        foreach (PersonUniversalMarker marker in person.UniversalMarker)
                        {
                            IList<int> result = GetUniversalDocumentMarker().GetCrossedElements(marker.ElementId, marker.SelectionLastelementId);
                            foreach (int index in result)
                            {
                                if (string.IsNullOrEmpty(SubdividedParagraphs[index]))
                                {
                                    SubdividedParagraphs[index] = person.Email;
                                }
                                else
                                    throw new ElementToPersonPairException();
                            }
                        }
                    }
                }

                string email = string.Empty;
                OpenXMLDocumentPart<SlideId> part = new OpenXMLDocumentPart<SlideId>();
                var slidePartsList = Presentation.Presentation.SlideIdList.ChildElements;
                for (int index = 0; index < slidePartsList.Count; index++)
                {
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<SlideId>();
                        part.CompositeElements.Add(slidePartsList[index] as SlideId);
                        email = SubdividedParagraphs[index];
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                        part.CompositeElements.Add(slidePartsList[index] as SlideId);
                }
            }

            return documentElements;
        }
	}

	public class PresentationSplit : MergeXml<SlideId>, ISplit, ILocalSplit
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
                IPresentationMarkerMapper mapping = new MarkerPresentationMapper(DocumentName, splitXml, body);
                DocumentElements = mapping.Run();
            }
        }

		List<PersonFiles> ISplit.SaveSplitDocument(Stream document)
		{
            List<PersonFiles> resultList = new List<PersonFiles>();

            byte[] byteArray = ReadFully(document);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);                
                using (PresentationDocument preDoc =
                    PresentationDocument.Open(mem, true))
                {
                    PresentationPart presentationPart = preDoc.PresentationPart;
                    //using (PresentationDocument templateDocument = PresentationDocument.Open(mem, false))
                    //{
                    //    foreach (SlideId slideId in templateDocument.PresentationPart.Presentation.SlideIdList.ChildElements)
                    //    {
                    //        presentationPart.DeletePart(slideId.RelationshipId);
                    //    }
                    //}
                    foreach (OpenXMLDocumentPart<SlideId> element in DocumentElements)
                    {
                        PresentationTools.RemoveAllSlides(presentationPart);

                        //alternative RemoveAllSlides
                        //presentationPart.Presentation = new Presentation();
                        //presentationPart.Presentation.SlideIdList = new SlideIdList();         

                        foreach (SlideId compo in element.CompositeElements)
                        {
                            PresentationTools.InsertSlideFromTemplate(presentationPart, mem, compo.RelationshipId);
                            //PresentationTools.InsertNewSlide(preDoc, 1, "aaaa");
                        }

                        presentationPart.Presentation.Save();

                        var person = new PersonFiles();
                        person.Person = element.PartOwner;
                        resultList.Add(person);
                        person.Name = element.Guid.ToString();
                        person.Data = mem.ToArray();
                    }
                }
            }
            // At this point, the memory stream contains the modified document.
            // We could write it back to a SharePoint document library or serve
            // it from a web server.			
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (PresentationDocument preDoc =
                    PresentationDocument.Open(mem, true))
                {
                    PresentationTools.RemoveAllSlides(preDoc.PresentationPart);
                    preDoc.PresentationPart.Presentation.Save();

                    var person = new PersonFiles();
                    person.Person = "/";
                    resultList.Add(person);
                    person.Name = "template.pptx";
                    person.Data = mem.ToArray();
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
            var docSplit = new DocumentSplit(DocumentName);

            return docSplit.CreateSplitXml(parts);
        }

        public List<PartsSelectionTreeElement> PartsFromSplitXml(Stream xmlFile, List<PartsSelectionTreeElement> parts)
        {
            var docSplit = new DocumentSplit(DocumentName);

            return docSplit.PartsFromSplitXml(xmlFile, parts);
        }

        #region Private methods
        #endregion
    }
}
