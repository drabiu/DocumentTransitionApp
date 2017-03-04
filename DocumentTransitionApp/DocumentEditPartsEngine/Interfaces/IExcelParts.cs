using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngine.Interfaces
{
    public interface IExcelParts
    {
        List<PartsSelectionTreeElement> Get(Stream file, Predicate<OpenXmlElement> supportedTypes);
        List<PartsSelectionTreeElement> GetSheets(Stream file);
    }
}
