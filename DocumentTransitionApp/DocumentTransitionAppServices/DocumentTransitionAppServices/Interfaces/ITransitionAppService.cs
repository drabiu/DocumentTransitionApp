using DocumentEditPartsEngine;
using SplitDescriptionObjects;
using System.Collections.Generic;

namespace DocumentTransitionAppServices.Interfaces
{
    public interface IDocumentService
    {
        PersonFiles[] SplitDocument(string docName, byte[] docxFile, byte[] xmlFile);
        byte[] GenerateSplitDocument(string docName, PartsSelectionTreeElement[] parts);
        List<PartsSelectionTreeElement> GetDocumentParts(string docName, byte[] documentFile);
        ServiceResponse GetDocumentPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);

    }

    public interface IPresentationService
    {
        PersonFiles[] SplitPresentation(string docName, byte[] docFile, byte[] xmlFile);
        byte[] GenerateSplitPresentation(string docName, PartsSelectionTreeElement[] parts);
        List<PartsSelectionTreeElement> GetPresentationParts(string preName, byte[] presentationFile);
        ServiceResponse GetPresentationPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);
    }

    public interface IExcelService
    {
        PersonFiles[] SplitExcel(string docName, byte[] docFile, byte[] xmlFile);
        byte[] GenerateSplitExcel(string docName, PartsSelectionTreeElement[] parts);
        List<PartsSelectionTreeElement> GetExcelParts(string excName, byte[] excelFile);
        ServiceResponse GetExcelPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);
    }

    public interface ITransitionAppService : IDocumentService, IPresentationService, IExcelService
    {
        byte[] MergeDocument(PersonFiles[] files);                           
    }
}