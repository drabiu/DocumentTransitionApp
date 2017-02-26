using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.IO;

namespace OpenXMLTools.Interfaces
{
    public interface IPresentationTools
    {
        PresentationDocument InsertSlideFromTemplate(PresentationDocument target, PresentationDocument template, IList<string> slideRelationshipIdList);

        PresentationDocument InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle);

        PresentationDocument RemoveAllSlides(PresentationDocument presentationDocument);
    }
}
