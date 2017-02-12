using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
