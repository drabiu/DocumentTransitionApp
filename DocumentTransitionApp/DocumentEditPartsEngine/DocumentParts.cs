using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
		public ElementType Type { get; private set; }
		public List<PartsSelectionTreeElement> Childs { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }

		public PartsSelectionTreeElement()
		{
		}

		public PartsSelectionTreeElement(string id, ElementType type, string name, int indent)
		{
			this.Id = id;
			this.Type = type;
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
		List<PartsSelectionTreeElement> IDocumentParts.Get(Stream file)
		{
			List<PartsSelectionTreeElement> documentElements = new List<PartsSelectionTreeElement>();
			using (WordprocessingDocument wordDoc =
				WordprocessingDocument.Open(file, true))
			{
				Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;			
				for (int index = 0; index < body.ChildElements.Count; index++)
				{
					PartsSelectionTreeElement part;
					var element = body.ChildElements[index];
					part = GetPicturePart(element);
					part = GetTablePart(element);
					if (part != null)
						documentElements.Add(part);
				}
			}

			return documentElements;
		}

		private PartsSelectionTreeElement GetPicturePart(OpenXmlElement element)
		{
			PartsSelectionTreeElement result = null;
			if (element is Wordproc.Paragraph)
			{
				Wordproc.Paragraph paragraph = element as Wordproc.Paragraph;
                if (paragraph.ChildElements.Any(ch => ch is Wordproc.Run))
				{
					Wordproc.Run run = paragraph.ChildElements.OfType<Wordproc.Run>().SingleOrDefault();
					if (run.ChildElements.Any(ch => ch is Wordproc.Picture))
					{
						result = new PartsSelectionTreeElement(paragraph.ParagraphId, ElementType.Picture, string.Empty, 0);
					}
				}
			}

			return result;
		}

		private PartsSelectionTreeElement GetTablePart(OpenXmlElement element)
		{
			PartsSelectionTreeElement result = null;
			if (element is Wordproc.Table)
			{

			}

			return result;
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
