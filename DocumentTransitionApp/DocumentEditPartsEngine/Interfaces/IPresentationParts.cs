using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngine.Interfaces
{
    public interface IPresentationParts
    {
        List<PartsSelectionTreeElement> GetSlidesWithAdditionalPats(Stream file, Predicate<OpenXmlPart> supportedParts);
        List<PartsSelectionTreeElement> GetSlides(Stream file);
    }
}
