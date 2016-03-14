using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
            var indexes = MarkerHelper<OpenXmlElement>.GetCrossedElements(id, id2, DocumentBody.ChildElements.ToList(), element => (element as Paragraph).ParagraphId.Value);

            return indexes;
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
