using DocumentEditPartsEngine;
using DocumentTransitionAppService;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.ServiceModel;

namespace DocumentTransitionAppWCF
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "ITransitionAppService" in both code and config file together.
    [ServiceContract]
    public interface ITransitionAppService
    {
        [OperationContract]
        PersonFiles[] SplitWord(string docName, byte[] docxFile, byte[] xmlFile);

        [OperationContract]
        PersonFiles[] SplitPresentation(string docName, byte[] docFile, byte[] xmlFile);

        [OperationContract]
        byte[] GenerateSplitWord(string docName, PartsSelectionTreeElement[] parts);

        [OperationContract]
        byte[] GenerateSplitPresentation(string docName, PartsSelectionTreeElement[] parts);

        [OperationContract]
        List<PartsSelectionTreeElement> GetWordParts(string docName, byte[] documentFile);

        [OperationContract]
        List<PartsSelectionTreeElement> GetPresentationParts(string preName, byte[] presentationFile);

        [OperationContract]
        ServiceResponse GetWordPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);

        [OperationContract]
        ServiceResponse GetPresentationPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);

        [OperationContract]
        PersonFiles[] SplitExcel(string docName, byte[] docFile, byte[] xmlFile);

        [OperationContract]
        byte[] GenerateSplitExcel(string docName, PartsSelectionTreeElement[] parts);

        [OperationContract]
        List<PartsSelectionTreeElement> GetExcelParts(string excName, byte[] excelFile);

        [OperationContract]
        ServiceResponse GetExcelPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);

        [OperationContract]
        byte[] MergeWord(PersonFiles[] files);

        [OperationContract]
        byte[] MergePresentation(PersonFiles[] files);

        [OperationContract]
        byte[] MergeExcel(PersonFiles[] files);
    }
}
