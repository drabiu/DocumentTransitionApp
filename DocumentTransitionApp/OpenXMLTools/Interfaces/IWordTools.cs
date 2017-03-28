using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTools.Interfaces
{
    public interface IWordTools
    {
        WordprocessingDocument MergeWordMedia(WordprocessingDocument target, WordprocessingDocument source);
        WordprocessingDocument MergeWordEmbeddings(WordprocessingDocument target, WordprocessingDocument source);
        WordprocessingDocument MergeWordCharts(WordprocessingDocument target, WordprocessingDocument source);
    }
}
