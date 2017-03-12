using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.IO;

namespace OpenXMLTools.Interfaces
{
    public interface IPresentationTools
    {
        PresentationDocument InsertSlidesFromTemplate(PresentationDocument target, PresentationDocument template, IList<string> slideRelationshipIdList);

        PresentationDocument InsertSlidesFromTemplate(PresentationDocument target, PresentationDocument template);

        PresentationDocument InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle);

        PresentationDocument RemoveAllSlides(PresentationDocument presentationDocument);

        PresentationDocument DeleteSlide(PresentationDocument presentationDocument, int slideIndex);
    }
}
