using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentEditPartsEngine.Interfaces;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

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
            result.Add(new PartsSelectionTreeElement(id.ToString(), elementId, GetSlideTitle(slidePart), 0));

            return result;
        }

        public static string GetSlideTitle(SlidePart slidePart)
        {
            if (slidePart == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            string paragraphSeparator = null;
            if (slidePart.Slide != null)
            {
                var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                             where IsTitleShape(shape)
                             select shape;

                StringBuilder paragraphText = new StringBuilder();
                foreach (var shape in shapes)
                {
                    foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
                    {
                        paragraphText.Append(paragraphSeparator);
                        paragraphText.Append("[Sld]: ");
                        foreach (var text in paragraph.Descendants<D.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        paragraphSeparator = "\n";
                    }
                }

                StringBuilder result = new StringBuilder();
                var listWords = paragraphText.ToString().Split(default(char[]), StringSplitOptions.RemoveEmptyEntries);
                foreach (var word in listWords)
                {
                    result.Append(string.Format("{0} ", word));
                    if (result.Length > PresentationDocumentPartAttributes.MaxNameLength)
                        break;
                }

                result.Remove(result.Length - 1, 1);

                return result.ToString();
            }

            return string.Empty;
        }

        private static bool IsTitleShape(Shape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    case PlaceholderValues.Title:
                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }

            return false;
        }      
    }
}
