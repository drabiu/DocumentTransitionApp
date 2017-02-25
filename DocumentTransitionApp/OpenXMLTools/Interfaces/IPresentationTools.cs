using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace OpenXMLTools.Interfaces
{
    public interface IPresentationTools
    {
        PresentationDocument InsertSlideFromTemplate(PresentationDocument target, PresentationDocument template, string sourceRelationshipId);

        PresentationDocument InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle);

        PresentationDocument RemoveAllSlides(PresentationDocument presentationDocument);
    }
}
