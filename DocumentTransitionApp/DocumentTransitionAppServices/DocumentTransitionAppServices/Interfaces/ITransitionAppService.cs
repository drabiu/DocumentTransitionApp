using DocumentEditPartsEngine;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DocumentTransitionAppServices.Interfaces
{
    public interface ITransitionAppService
    {
        PersonFiles[] SplitDocument(string docName, byte[] docxFile, byte[] xmlFile);
        PersonFiles[] SplitPresentation(string docName, byte[] docFile, byte[] xmlFile);
        byte[] GenerateSplitDocument(string docName, PartsSelectionTreeElement[] parts);
        byte[] GenerateSplitPresentation(string docName, PartsSelectionTreeElement[] parts);
        byte[] MergeDocument(PersonFiles[] files);
        List<PartsSelectionTreeElement> GetDocumentParts(string docName, byte[] documentFile);
        List<PartsSelectionTreeElement> GetPresentationParts(string preName, byte[] presentationFile);
        ServiceResponse GetDocumentPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);
        ServiceResponse GetPresentationPartsFromXml(string docName, byte[] documentFile, byte[] splitFile);
    }
}