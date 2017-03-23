using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentEditPartsEngine.Interfaces
{
    public interface IDocumentParts
    {
        List<PartsSelectionTreeElement> Get(Stream file, Predicate<OpenXmlElement> supportedType);
        List<PartsSelectionTreeElement> Get(Stream file);
        PartsSelectionTreeElement GetParagraphSelectionTreeElement(OpenXmlElement element, PartsSelectionTreeElement parent, int id, Predicate<OpenXmlElement> supportedType, int indent);
    }
}
