using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
			throw new NotImplementedException();
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
