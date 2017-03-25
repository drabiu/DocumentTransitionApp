using System;
using System.Collections.Generic;

namespace DocumentSplitEngine.Data_Structures
{
    public class OpenXMLDocumentPart<Element>
    {
        public IList<Element> CompositeElements { get; set; }
        public string PartOwner { get; set; }
        public Guid Guid { get; private set; }

        public OpenXMLDocumentPart()
        {
            this.Guid = Guid.NewGuid();
            CompositeElements = new List<Element>();
        }
    }
}
