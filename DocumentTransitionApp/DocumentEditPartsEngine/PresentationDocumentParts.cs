﻿using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OpenXMLTools;
using OpenXMLTools.Helpers;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngine
{
    public static class PresentationDocumentPartAttributes
    {
        public const int MaxNameLength = 36;

        public static bool IsSupportedPart(OpenXmlPart part)
        {
            bool isSupported = false;
            isSupported = part is SlidePart;

            return isSupported;
        }
    }

    public class PresentationDocumentParts : IPresentationParts
    {
        public List<PartsSelectionTreeElement> Get(Stream file, Predicate<OpenXmlPart> supportedParts)
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
                    presentationElements.AddRange(CreatePartsSelectionTreeElements(slidePart, idIndex, elementId, supportedParts, 0));
                    idIndex++;
                }
            }

            return presentationElements;
        }

        public List<PartsSelectionTreeElement> GetSlides(Stream file)
        {
            return Get(file, el => PresentationDocumentPartAttributes.IsSupportedPart(el));
        }

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlPart openXmlPart, int id, string elementId, Predicate<OpenXmlPart> isSupportedPart, int indent)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (isSupportedPart(openXmlPart))
            {
                if (openXmlPart is SlidePart)
                {
                    result.Add(new PartsSelectionTreeElement(id.ToString(), elementId, PresentationTools.GetSlideTitle(openXmlPart as SlidePart, PresentationDocumentPartAttributes.MaxNameLength), indent, ElementType.Slide));
                }
            }

            return result;
        }
    }
}
