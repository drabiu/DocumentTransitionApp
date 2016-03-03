using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SplitDescriptionObjects
{
	public interface IDocumentMarker
	{
		int FindElement(string id);
		IList<int> GetCrossedElements(string id, string id2);
	}

	public abstract class DocumentMarker : IDocumentMarker
	{
		Body DocumentBody;

		public DocumentMarker(Body body)
		{
			DocumentBody = body;
		}

		public int FindElement(string id)
		{
			throw new NotImplementedException();
		}

		public IList<int> GetCrossedElements(string id, string id2)
		{
			bool startSelection = false;
			IList<int> indexes = new List<int>();
			for (int index = 0; index < DocumentBody.ChildElements.Count; index++)
			{
				OpenXmlElement element = DocumentBody.ChildElements[index];
				if (element is Paragraph)
				{
					if ((element as Paragraph).ParagraphId.Value == id)
						startSelection = true;

                    if (startSelection)
                        indexes.Add(index);

                    if ((element as Paragraph).ParagraphId.Value == id2)
						break;
				}
			}

			return indexes;
		}
	}

	public abstract class ExcelMarker : IDocumentMarker
	{
		Workbook DocumentBody;

		public ExcelMarker(Workbook body)
		{
			DocumentBody = body;
		}

		public int FindElement(string id)
		{
			throw new NotImplementedException();
		}

		public IList<int> GetCrossedElements(string id, string id2)
		{
			throw new NotImplementedException();
		}
	}

	public interface IUniversalDocumentMarker : IDocumentMarker
	{
	}

	public class UniversalDocumentMarker : DocumentMarker, IUniversalDocumentMarker
	{
		public UniversalDocumentMarker(Body body) :
			base(body)
		{
		}
	}

	public interface ISheetExcelMarker : IDocumentMarker
	{
	}

	public class SheetExcelMarker : ExcelMarker, ISheetExcelMarker
	{
		public SheetExcelMarker(Workbook body) :
			base(body)
		{

		}
	}
}
