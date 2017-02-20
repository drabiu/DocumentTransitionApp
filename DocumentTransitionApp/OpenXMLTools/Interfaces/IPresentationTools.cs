using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace OpenXMLTools.Interfaces
{
    public interface IPresentationTools
    {
        void InsertSlideFromTemplate(PresentationPart presentationPart, MemoryStream mem, string sourceRelationshipId);

        void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle);

        void RemoveAllSlides(PresentationPart presentationPart);
    }
}
