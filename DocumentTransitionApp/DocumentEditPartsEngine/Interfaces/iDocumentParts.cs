using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngine.Interfaces
{
    public interface IDocumentParts
    {
        List<PartsSelectionTreeElement> Get(Stream file, Predicate<OpenXmlElement> supportedType);
        List<PartsSelectionTreeElement> Get(Stream file);
        PartsSelectionTreeElement GetParagraphSelectionTreeElement(OpenXmlElement element, PartsSelectionTreeElement parent, int id, Predicate<OpenXmlElement> supportedType, int indent, bool visible);
    }
}
