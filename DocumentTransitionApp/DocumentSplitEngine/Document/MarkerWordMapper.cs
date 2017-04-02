using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentSplitEngine.Data_Structures;
using DocumentSplitEngine.Interfaces;
using SplitDescriptionObjects;
using SplitDescriptionObjects.DocumentMarkers;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace DocumentSplitEngine.Document
{
    public class MarkerWordMapper : MarkerDocumentMapper, IMarkerMapper<OpenXmlElement>
    {
        SplitDocument SplitDocumentObj { get; set; }
        Body DocumentBody { get; set; }
        List<MarkerWordSelector> _markerWordSelectors;

        public MarkerWordMapper(string documentName, Split xml, Body body)
        {
            Xml = xml;
            SplitDocumentObj = (SplitDocument)Xml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, documentName)).SingleOrDefault();
            DocumentBody = body;
            _markerWordSelectors = MarkerWordSelector.InitializeSelectorsList(DocumentBody);
        }

        public IList<OpenXMLDocumentPart<OpenXmlElement>> Run()
        {
            IUniversalWordMarker universalDocMarker;
            IPictureWordMarker pictureDocMarker;
            IListWordMarker listDocMarker;
            ITableWordMarker tableDocMarker;

            IList<OpenXMLDocumentPart<OpenXmlElement>> documentElements = new List<OpenXMLDocumentPart<OpenXmlElement>>();
            if (SplitDocumentObj?.Person != null)
            {
                foreach (Person person in SplitDocumentObj.Person)
                {
                    universalDocMarker = new UniversalWordMarker(DocumentBody, _markerWordSelectors);
                    _markerWordSelectors = universalDocMarker.GetSubdividedParts(person);

                    listDocMarker = new ListWordMarker(DocumentBody, _markerWordSelectors);
                    _markerWordSelectors = listDocMarker.GetSubdividedParts(person);

                    pictureDocMarker = new PictureWordMarker(DocumentBody, _markerWordSelectors);
                    _markerWordSelectors = pictureDocMarker.GetSubdividedParts(person);

                    tableDocMarker = new TableWordMarker(DocumentBody, _markerWordSelectors);
                    _markerWordSelectors = tableDocMarker.GetSubdividedParts(person);
                }

                string email = string.Empty;
                OpenXMLDocumentPart<OpenXmlElement> part = new OpenXMLDocumentPart<OpenXmlElement>();
                foreach (var wordSelector in _markerWordSelectors)
                {
                    if (wordSelector.Email != email)
                    {
                        part = new OpenXMLDocumentPart<OpenXmlElement>();
                        part.CompositeElements.Add(wordSelector.Element);
                        email = wordSelector.Email;
                        if (string.IsNullOrEmpty(email))
                            part.PartOwner = "undefined";
                        else
                            part.PartOwner = email;

                        documentElements.Add(part);
                    }
                    else
                        part.CompositeElements.Add(wordSelector.Element);
                }
            }

            return documentElements;
        }
    }
}
