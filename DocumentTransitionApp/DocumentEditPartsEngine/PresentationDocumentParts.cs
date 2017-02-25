using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using OpenXMLTools;

namespace DocumentEditPartsEngine
{
    public static class PresentationDocumentPartAttributes
    {
        public const int MaxNameLength = 36;

        public static bool IsSupportedPart(OpenXmlPart part)
        {
            bool isSupported = false;
            isSupported = part is SlidePart;
            //|| element is Wordproc.Picture
            //|| element is Wordproc.Drawing
            //|| element is Wordproc.Table;

            return isSupported;
        }
    }

    public class PresentationDocumentParts : IPresentationParts
	{
        public List<PartsSelectionTreeElement> GetSlidesWithAdditionalPats(Stream file, Predicate<OpenXmlPart> supportedParts)
        {
            throw new NotImplementedException();
        }

        public List<PartsSelectionTreeElement> GetSlides(Stream file)
		{
            List<PartsSelectionTreeElement> presentationElements = new List<PartsSelectionTreeElement>();
            using (PresentationDocument preDoc =
                PresentationDocument.Open(file, true))
            {
                Presentation presentation = preDoc.PresentationPart.Presentation;
                var idIndex = 1;
                foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = preDoc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                    string elementId = slideId.RelationshipId;
                    presentationElements.AddRange(CreatePartsSelectionTreeElements(slidePart, idIndex, elementId));
                    idIndex++;
                }
            }

            return presentationElements;
        }

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(SlidePart slidePart, int id, string elementId)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            result.Add(new PartsSelectionTreeElement(id.ToString(), elementId, PresentationTools.GetSlideTitle(slidePart, PresentationDocumentPartAttributes.MaxNameLength), 0));

            return result;
        }        
    }
}
