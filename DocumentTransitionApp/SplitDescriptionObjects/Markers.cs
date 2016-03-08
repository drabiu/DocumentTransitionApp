using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SplitDescriptionObjects
{
    public static class MarkerHelper<ElementType>
    {
        public static IList<int> GetCrossedElements(string id, string id2, IList<ElementType> elements, Func<ElementType, string> getElementId)
        {
            bool startSelection = false;
            IList<int> indexes = new List<int>();
            for (int index = 0; index < elements.Count; index++)
            {
                var element = elements[index];
                if (element is ElementType)
                {
                    if (getElementId(element) == id)
                        startSelection = true;

                    if (startSelection)
                        indexes.Add(index);

                    if (getElementId(element) == id2)
                        break;
                }
            }

            return indexes;
        }
    }
    
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
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, DocumentBody.ChildElements.ToList(), element => (element as Paragraph).ParagraphId.Value);

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

        //public static IEnumerable<Person> GetPersonsFromSplitXml(string documentName, Split splitXml)
        //{
        //    List<Person> persons = new List<Person>();
        //    var splitDocument = (SplitDocument)splitXml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, documentName)).SingleOrDefault();
        //    foreach (var person in splitDocument.Person)
        //    {
        //        var personObj = new Person() { Email = person.Email };
        //        personObj.UniversalMarker = new PersonUniversalMarker[splitDocument.Person.Count()];
        //        foreach (var universalMarker in person.UniversalMarker)
        //        {
        //            personObj.UniversalMarker = new PersonUniversalMarker() { ElementId = universalMarker.ElementId, SelectionLastelementId = universalMarker.SelectionLastelementId };
        //            markers.Add(new PersonUniversalMarker() { ElementId = universalMarker.ElementId })
        //        }
        //    }
        //}
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
