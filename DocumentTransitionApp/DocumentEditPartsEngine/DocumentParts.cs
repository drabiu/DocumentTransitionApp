using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;
using Present = DocumentFormat.OpenXml.Presentation;

namespace DocumentEditPartsEngine
{
	public enum ElementType
	{
		Paragraph,
		Table,
		Picture,
		Sheet,
		Slide
	}

	public class PartsSelectionTreeElement
	{
		public string Id { get; set; }
        public string ElementId { get; set; }
		//public ElementType Type { get; private set; }
		public List<PartsSelectionTreeElement> Childs { get; set; }
		public string Name { get; set; }
		public int Indent { get; set; }
        public string OwnerName { get; set; }
        public bool Selected { get; set; }

        public PartsSelectionTreeElement()
		{
		}

		public PartsSelectionTreeElement(string id, string name, int indent)
		{
			this.Id = id;
			//this.Type = type;
			this.Name = name;
			this.Indent = indent;
			this.Childs = new List<PartsSelectionTreeElement>();
		}

        public PartsSelectionTreeElement(string id, string elementId, string name, int indent) : this(id, name, indent)
        {
            this.ElementId = elementId;
        }

    }

	public interface IDocumentParts
	{
		List<PartsSelectionTreeElement> Get(Stream file);
	}

	public class DocumentPartsBuilder
	{
		public static IDocumentParts Build(string fileExtension)
		{
			IDocumentParts result;
			switch (fileExtension)
			{
				case (".docx"):
					result = new WordDocumentParts();
					break;
				case (".xlsx"):
					result = new ExcelDocumentParts();
					break;
				case (".pptx"):
					result = new PresentationDocumentParts();
					break;
				default:
					result = new WordDocumentParts();
					break;
			}

			return result;
		}
	}

	public class WordDocumentParts : IDocumentParts
	{
		private class WordDocumentPartAttributes
		{
			public const int MaxNameLength = 30;
		}

		List<PartsSelectionTreeElement> IDocumentParts.Get(Stream file)
		{
			List<PartsSelectionTreeElement> documentElements = new List<PartsSelectionTreeElement>();
			using (WordprocessingDocument wordDoc =
				WordprocessingDocument.Open(file, true))
			{
				Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
				for (int index = 0; index < body.ChildElements.Count; index++)
				{
					var element = body.ChildElements[index];
					documentElements.AddRange(CreatePartsSelectionTreeElements(element, index));
				}
			}

			return documentElements;
		}

        private IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, int id)
		{
			List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
			if (IsSupportedType(element))
			{
                PartsSelectionTreeElement elementToAdd;
                if (element is Wordproc.Paragraph)
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), (element as Wordproc.Paragraph).ParagraphId, GetElementName(element), 0);
                else
                    elementToAdd = new PartsSelectionTreeElement(id.ToString(), GetElementName(element), 0);

                result.Add(elementToAdd);
				if(element.HasChildren)
				{
					CreateChildrenPartsSelectionTreeElements(element);
                }
			}

			return result;
		}

		private IEnumerable<PartsSelectionTreeElement> CreateChildrenPartsSelectionTreeElements(OpenXmlElement element)
		{
			List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
			return result;
		}

		private string GetElementName(OpenXmlElement element)
		{
			StringBuilder result = new StringBuilder();
			if (element is Wordproc.Paragraph)
			{
				var paragraph = element as Wordproc.Paragraph;
				if (paragraph.ChildElements.Any(ch => ch is Wordproc.Run))
				{
					result.Append("Paragraph: ");
					foreach (Wordproc.Run run in paragraph.ChildElements.OfType<Wordproc.Run>())
					{
						result.Append(run.InnerText);
						if (result.Length > WordDocumentPartAttributes.MaxNameLength)
							break;
                    }			
				}
			}
			else if (element is Wordproc.Table)
			{


			}
			else if (element is Wordproc.Picture)
			{

			}
			else if (element is Wordproc.Drawing)
			{ }

			return result.ToString();
		}

		private bool IsSupportedType(OpenXmlElement element)
		{
			bool isSupported = false;
			isSupported = element is Wordproc.Paragraph;
				//|| element is Wordproc.Picture
				//|| element is Wordproc.Drawing
				//|| element is Wordproc.Table;

			return isSupported;
		}
    }

	public class ExcelDocumentParts : IDocumentParts
	{
		public List<PartsSelectionTreeElement> Get(Stream file)
		{
			throw new NotImplementedException();
		}
    }

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
                var index = 1;
                foreach (var slideId in presentation.SlideIdList.Elements<Present.SlideId>())
                {
                    SlidePart slidePart = preDoc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                    presentationElements.AddRange(CreatePartsSelectionTreeElements(slidePart, index++, slideId.RelationshipId));
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
