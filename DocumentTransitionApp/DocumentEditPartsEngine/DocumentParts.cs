using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;

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
		public string Id { get; private set; }
		//public ElementType Type { get; private set; }
		public List<PartsSelectionTreeElement> Childs { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }

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
				result.Add(new PartsSelectionTreeElement(id.ToString(), GetElementName(element), 0));
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
		public List<PartsSelectionTreeElement> Get(Stream file)
		{
			throw new NotImplementedException();
		}
	}
}
