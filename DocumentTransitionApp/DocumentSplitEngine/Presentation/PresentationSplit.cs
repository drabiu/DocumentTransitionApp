﻿using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentSplitEngine.Data_Structures;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Presentproc = DocumentFormat.OpenXml.Presentation;

namespace DocumentSplitEngine
{
    public interface IPresentationMarkerMapper
    {
        IList<OpenXMLDocumentPart<Presentproc.SlideId>> Run();
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
        public IList<OpenXMLDocumentPart<Presentproc.SlideId>> Run()
		{
            IList<OpenXMLDocumentPart<Presentproc.SlideId>> documentElements = new List<OpenXMLDocumentPart<Presentproc.SlideId>>();
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
                OpenXMLDocumentPart<Presentproc.SlideId> part = new OpenXMLDocumentPart<Presentproc.SlideId>();
                var slidePartsList = Presentation.Presentation.SlideIdList.ChildElements;
                for (int index = 0; index < slidePartsList.Count; index++)
                {
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<Presentproc.SlideId>();
                        part.CompositeElements.Add(slidePartsList[index] as Presentproc.SlideId);
                        email = SubdividedParagraphs[index];
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                        part.CompositeElements.Add(slidePartsList[index] as Presentproc.SlideId);
                }
            }

            return documentElements;
        }
	}

	public class PresentationSplit : MergeXml<Presentproc.SlideId>, ISplit, ILocalSplit
	{
        public PresentationSplit(string docName)
        {
            DocumentName = docName;
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
                    RemoveAllSlides(presentationPart);
                    foreach (OpenXMLDocumentPart<Presentproc.SlideId> element in DocumentElements)
                    {
                        var slideIdList = presentationPart.Presentation.SlideIdList;
                        //Find the highest slide ID in the current list.
                        uint maxSlideId = 1;

                        foreach (Presentproc.SlideId slideId in slideIdList.ChildElements)
                        {
                            if (slideId.Id.Value > maxSlideId)
                            {
                                maxSlideId = slideId.Id;
                            }
                        }

                        maxSlideId ++;

                        foreach (Presentproc.SlideId compo in element.CompositeElements)
                        {
                            PresentationDocument templateDocument = PresentationDocument.Open(mem, false);
                            SlidePart templateSlide = (SlidePart)templateDocument.PresentationPart.GetPartById(compo.RelationshipId);
                            //Create the slide part and copy the data from the first part
                            SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();
                            newSlidePart.FeedData(templateSlide.GetStream());
                            //Use the same slide layout as that of the template slide.
                            if (null != templateSlide.SlideLayoutPart)
                            {
                                newSlidePart.AddPart(templateSlide.SlideLayoutPart);
                            }

                            templateDocument.Close();

                            //Insert the new slide into the slide list.
                            Presentproc.SlideId newSlideId = slideIdList.AppendChild<Presentproc.SlideId>(new Presentproc.SlideId());
                            //Presentproc.SlideId newSlideId = slideIdList.InsertAfter(new Presentproc.SlideId(), slideIdList.Last());

                            //Set the slide id and relationship id
                            newSlideId.Id = maxSlideId;
                            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);
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
