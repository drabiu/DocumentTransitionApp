using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentEditPartsEngine;

namespace SplitDescriptionObjects
{
    public interface IDocumentMarker
    {
        int FindElement(string id);
        IList<int> GetCrossedParagraphElements(string id, string id2);
    }

    public abstract class DocumentMarker : IDocumentMarker
    {
        Body DocumentBody;
        List<OpenXmlElement> ElementsList;

        public DocumentMarker(Body body)
        {
            DocumentBody = body;
            ElementsList = DocumentBody.ChildElements.ToList();
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        public IList<int> GetCrossedParagraphElements(string id, string id2)
        {
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, ElementsList, el => el is Paragraph, element => GetParagraphId(element));

            return indexes;
        }

        private string GetParagraphId(OpenXmlElement element)
        {
            string result = string.Empty;
            if (element is Paragraph)
            {
                Paragraph parahraph = (element as Paragraph);
                int index = ElementsList.FindIndex(el => el.Equals(element));
                result = parahraph.ParagraphId != null ? parahraph.ParagraphId.Value : WordDocumentPartAttributes.GetParagraphNoIdFormatter(index);
            }        

            return result;
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
}
