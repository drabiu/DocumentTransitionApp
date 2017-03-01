using DocumentEditPartsEngine;
using System.Collections.Generic;
using System.IO;

namespace DocumentSplitEngine.Interfaces
{
    public interface ISplitXml
    {
        byte[] CreateSplitXml(IList<PartsSelectionTreeElement> parts);
        List<PartsSelectionTreeElement> SelectPartsFromSplitXml(Stream xmlFile, List<PartsSelectionTreeElement> parts);
    }
}
