using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using SplitDescriptionObjects;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace DocumentSplitEngine.Document
{
    public class MarkerWordMapper : MarkerDocumentMapper, IMarkerMapper<OpenXmlElement>
    {
        SplitDocument SplitDocumentObj { get; set; }
        Body DocumentBody { get; set; }
        IUniversalWordMarker UniversalDocMarker;

        public MarkerWordMapper(string documentName, Split xml, Body body)
        {
            Xml = xml;
            SplitDocumentObj = (SplitDocument)Xml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, documentName)).SingleOrDefault();
            DocumentBody = body;
            UniversalDocMarker = new UniversalWordMarker(DocumentBody);
            SubdividedParagraphs = new string[body.ChildElements.Count];
        }

        public IList<OpenXMLDocumentPart<OpenXmlElement>> Run()
        {
            IList<OpenXMLDocumentPart<OpenXmlElement>> documentElements = new List<OpenXMLDocumentPart<OpenXmlElement>>();
            if (SplitDocumentObj?.Person != null)
            {
                foreach (Person person in SplitDocumentObj.Person)
                {
                    //if (person.UniversalMarker != null)
                    //{
                    //    foreach (PersonUniversalMarker marker in person.UniversalMarker)
                    //    {
                    //        IList<int> result = UniversalDocMarker.GetCrossedParagraphElements(marker.ElementId, marker.SelectionLastelementId);
                    //        foreach (int index in result)
                    //        {
                    //            if (string.IsNullOrEmpty(SubdividedParagraphs[index]))
                    //            {
                    //                SubdividedParagraphs[index] = person.Email;
                    //            }
                    //            else
                    //                throw new ElementToPersonPairException();
                    //        }
                    //    }
                    //}

                    //if (person.ListMarker != null)
                    //{
                    //}

                    //if (person.PictureMarker != null)
                    //{
                    //}

                    //if (person.TableMarker != null)
                    //{
                    //}

                    //UniversalWordMarker.SetPartsOwner(parts, person);
                    //TableWordMarker.SetPartsOwner(parts, person);
                    //ListWordMarker.SetPartsOwner(parts, person);
                    //PictureWordMarker.SetPartsOwner(parts, person);
                }

                string email = string.Empty;
                OpenXMLDocumentPart<OpenXmlElement> part = new OpenXMLDocumentPart<OpenXmlElement>();
                for (int index = 0; index < DocumentBody.ChildElements.Count; index++)
                {
                    //check if parts are neighbours then join into one document
                    if (SubdividedParagraphs[index] != email)
                    {
                        part = new OpenXMLDocumentPart<OpenXmlElement>();
                        part.CompositeElements.Add(DocumentBody.ChildElements[index]);
                        email = SubdividedParagraphs[index];
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                        part.CompositeElements.Add(DocumentBody.ChildElements[index]);
                }
            }

            return documentElements;
        }
    }
}
