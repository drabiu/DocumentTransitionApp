using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SplitDescriptionObjects
{
	interface IDocumentMarker
	{
		OpenXmlCompositeElement FindElement(string id);
		IList<OpenXmlCompositeElement> GetCrossedElements(string id, string id2);
	}

	public class DocumentMarker : IDocumentMarker
	{
		Body DocumentBody;

		public DocumentMarker(Body body)
		{
			DocumentBody = body;
		}

		public OpenXmlCompositeElement FindElement(string id)
		{
			throw new NotImplementedException();
		}

		public IList<OpenXmlCompositeElement> GetCrossedElements(string id, string id2)
		{
			throw new NotImplementedException();
		}
	}
}
