using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.Linq;

namespace DocumentSplitEngine.Presentation
{
    public class MarkerPresentationMapper : MarkerDocumentMapper, IMarkerMapper<SlideId>
    {
        SplitPresentation SplitPresentationObj { get; set; }
        PresentationPart Presentation;
        IUniversalPresentationMarker UniversalPreMarker;

        public MarkerPresentationMapper(string documentName, Split xml, PresentationPart presentation)
        {
            Xml = xml;            
            SplitPresentationObj = (SplitPresentation)Xml.Items.Where(it => it is SplitPresentation && string.Equals(((SplitPresentation)it).Name, documentName)).SingleOrDefault();
            Presentation = presentation;
            UniversalPreMarker = new UniversalPresentationMarker(Presentation);
            SubdividedParagraphs = new string[presentation.SlideParts.Count()];
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
                            IList<int> result = UniversalPreMarker.GetCrossedSlideIdElements(marker.ElementId, marker.SelectionLastelementId);
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
}
