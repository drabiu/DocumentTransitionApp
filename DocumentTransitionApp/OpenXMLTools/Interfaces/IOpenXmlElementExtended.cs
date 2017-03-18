using OpenXMLTools.Helpers;

namespace OpenXMLTools.Interfaces
{
    public interface IOpenXmlElementExtended
    {
        string GetElementName(int nameLength);
        ElementType GetElementType();
    }
}
