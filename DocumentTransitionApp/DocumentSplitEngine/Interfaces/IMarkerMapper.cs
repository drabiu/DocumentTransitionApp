using DocumentSplitEngine.Data_Structures;
using System.Collections.Generic;

namespace DocumentSplitEngine.Interfaces
{
    public interface IMarkerMapper<Element>
    {
        IList<OpenXMLDocumentPart<Element>> Run();
    }
}
