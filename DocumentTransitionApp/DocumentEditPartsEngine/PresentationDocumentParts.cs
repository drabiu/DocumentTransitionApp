using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentEditPartsEngine.Interfaces;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;
using Present = DocumentFormat.OpenXml.Presentation;

namespace DocumentEditPartsEngine
{
	public class PresentationDocumentParts : IDocumentParts
	{
        private class PresentationDocumentPartAttributes
        {
            public const int MaxNameLength = 30;
        }

        public List<PartsSelectionTreeElement> Get(Stream file)
		{
            List<PartsSelectionTreeElement> presentationElements = new List<PartsSelectionTreeElement>();
            using (PresentationDocument preDoc =
                PresentationDocument.Open(file, true))
            {
                Present.Presentation presentation = preDoc.PresentationPart.Presentation;
                var idIndex = 1;
                foreach (var slideId in presentation.SlideIdList.Elements<Present.SlideId>())
                {
                    SlidePart slidePart = preDoc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                    presentationElements.AddRange(CreatePartsSelectionTreeElements(slidePart, idIndex, (idIndex-1).ToString()));
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
                var shapes = from shape in slidePart.Slide.Descendants<Present.Shape>()
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

                return paragraphText.ToString();
            }

            return string.Empty;
        }

        private static bool IsTitleShape(Present.Shape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<Present.PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((Present.PlaceholderValues)placeholderShape.Type)
                {
                    case Present.PlaceholderValues.Title:
                    case Present.PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }

            return false;
        }
    }
}
