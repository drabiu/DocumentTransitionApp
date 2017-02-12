﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.IO;

using DocumentFormat.OpenXml.Packaging;
using Presentproc = DocumentFormat.OpenXml.Presentation;

using SplitDescriptionObjects;
using DocumentEditPartsEngine;
using DocumentSplitEngine.Data_Structures;

namespace DocumentSplitEngine
{ 
    public interface IPresentationMarkerMapper
    {
        IList<OpenXMLDocumentPart<SlidePart>> Run();
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

        public IList<OpenXMLDocumentPart<SlidePart>> Run()
		{
            IList<OpenXMLDocumentPart<SlidePart>> documentElements = new List<OpenXMLDocumentPart<SlidePart>>();
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
                OpenXMLDocumentPart<SlidePart> part = new OpenXMLDocumentPart<SlidePart>();
                var slidePartsList = Presentation.SlideParts.ToList();
                for (int index = 0; index < slidePartsList.Count; index++)
                {
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<SlidePart>();
                        part.CompositeElements.Add(slidePartsList[index]);
                        email = SubdividedParagraphs[index];
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                        part.CompositeElements.Add(slidePartsList[index]);
                }
            }

            return documentElements;
        }
	}

	public class PresentationSplit : MergeXml<SlidePart>, ISplit, ILocalSplit
	{
        public PresentationSplit(string docName)
        {
            DocumentName = docName;
        }

        public void OpenAndSearchDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
            throw new NotImplementedException();
        }

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
                PresentationDocument templateDocument = PresentationDocument.Open(mem, false);
                using (PresentationDocument preDoc =
                    PresentationDocument.Open(mem, true))
                {
                    //preDoc.DeletePart(preDoc.PresentationPart);
                    //PresentationPart presentationPart = preDoc.AddPresentationPart();
                    //presentationPart.Presentation = new Presentproc.Presentation();
                    PresentationPart presentationPart = preDoc.PresentationPart;
                    foreach (OpenXMLDocumentPart<SlidePart> element in DocumentElements)
                    {
                        RemoveAllSlides(presentationPart);
                        //foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<Presentproc.SlideId>())
                        //{
                        //    SlidePart slidePart = templateDocument.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        //}
                        foreach (SlidePart compo in element.CompositeElements)
                        {
                            //SlidePart slidePart = templateDocument.PresentationPart.GetPartById(compo.Slide.)
                            //presentationPart.AddPart<SlidePart>(compo);
                        }

                        presentationPart.Presentation.Save();

                        var person = new PersonFiles();
                        person.Person = element.PartOwner;
                        resultList.Add(person);
                        person.Name = element.Guid.ToString();
                        person.Data = mem.ToArray();
                    }
                }

                templateDocument.Close();
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
                    RemoveAllSlides(preDoc.PresentationPart);
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

        private void RemoveAllSlides(PresentationPart presentationPart)
        {
            Presentproc.Presentation presentation = presentationPart.Presentation;
            Presentproc. SlideIdList slideIdList = presentation.SlideIdList;

            foreach (Presentproc.SlideId slideId in slideIdList.ChildElements.ToList())
            {
                slideIdList.RemoveChild(slideId);
                string slideRelId = slideId.RelationshipId;

                if (presentation.CustomShowList != null)
                {
                    // Iterate through the list of custom shows.
                    foreach (var customShow in presentation.CustomShowList.Elements<Presentproc.CustomShow>())
                    {
                        if (customShow.SlideList != null)
                        {
                            // Declare a link list of slide list entries.
                            LinkedList<Presentproc.SlideListEntry> slideListEntries = new LinkedList<Presentproc.SlideListEntry>();
                            foreach (Presentproc.SlideListEntry slideListEntry in customShow.SlideList.Elements())
                            {
                                // Find the slide reference to remove from the custom show.
                                if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                {
                                    slideListEntries.AddLast(slideListEntry);
                                }
                            }

                            // Remove all references to the slide from the custom show.
                            foreach (Presentproc.SlideListEntry slideListEntry in slideListEntries)
                            {
                                customShow.SlideList.RemoveChild(slideListEntry);
                            }
                        }
                    }
                }
            }
        }
            
    }
}
