using DocumentEditPartsEngine;
using DocumentEditPartsEngine.Interfaces;
using DocumentMergeEngine;
using DocumentMergeEngine.Interfaces;
using DocumentSplitEngine;
using DocumentSplitEngine.Interfaces;
using DocumentTransitionAppWCF.Responses;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;

namespace DocumentTransitionAppWCF
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "TransitionAppService" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select TransitionAppService.svc or TransitionAppService.svc.cs at the Solution Explorer and start debugging.
    public class TransitionAppService : ITransitionAppService
    {
        public byte[] GenerateSplitExcel(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplitXml split = new ExcelSplit(docName);

            return split.CreateSplitXml(parts);
        }

        public byte[] GenerateSplitPresentation(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplitXml split = new PresentationSplit(docName);

            return split.CreateSplitXml(parts);
        }

        public byte[] GenerateSplitWord(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplitXml split = new WordSplit(docName);

            return split.CreateSplitXml(parts);
        }

        public List<PartsSelectionTreeElement> GetExcelParts(string excName, byte[] excelFile)
        {
            IDocumentParts parts = new ExcelDocumentParts();

            return parts.Get(new MemoryStream(excelFile));
        }

        public GetPartsFromXmlServiceResponse GetExcelPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            ISplitXml split = new ExcelSplit(Path.GetFileNameWithoutExtension(docName));
            var cleanParts = GetExcelParts(docName, documentFile);

            try
            {
                return new GetPartsFromXmlServiceResponse(split.SelectPartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
            }
            catch (SplitNameDifferenceExcception ex)
            {
                return new GetPartsFromXmlServiceResponse(ex.Message);
            }
        }

        public List<PartsSelectionTreeElement> GetPresentationParts(string preName, byte[] presentationFile)
        {
            IPresentationParts parts = new PresentationDocumentParts();

            return parts.GetSlides(new MemoryStream(presentationFile));
        }

        public GetPartsFromXmlServiceResponse GetPresentationPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            ISplitXml split = new PresentationSplit(Path.GetFileNameWithoutExtension(docName));
            var cleanParts = GetPresentationParts(docName, documentFile);

            try
            {
                return new GetPartsFromXmlServiceResponse(split.SelectPartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
            }
            catch (SplitNameDifferenceExcception ex)
            {
                return new GetPartsFromXmlServiceResponse(ex.Message);
            }
        }

        public List<PartsSelectionTreeElement> GetWordParts(string docName, byte[] documentFile)
        {
            IDocumentParts parts = new WordDocumentParts();

            return parts.Get(new MemoryStream(documentFile));
        }

        public GetPartsFromXmlServiceResponse GetWordPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            ISplitXml split = new WordSplit(Path.GetFileNameWithoutExtension(docName));
            var cleanParts = GetWordParts(docName, documentFile);

            try
            {
                return new GetPartsFromXmlServiceResponse(split.SelectPartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
            }
            catch (SplitNameDifferenceExcception ex)
            {
                return new GetPartsFromXmlServiceResponse(ex.Message);
            }
        }

        public byte[] MergeExcel(PersonFiles[] files)
        {
            IMerge merge = new ExcelMerge();

            return merge.Run(new List<PersonFiles>(files));
        }

        public byte[] MergePresentation(PersonFiles[] files)
        {
            IMerge merge = new PresentationMerge();

            return merge.Run(new List<PersonFiles>(files));
        }

        public byte[] MergeWord(PersonFiles[] files)
        {
            IMerge merge = new WordMerge();

            return merge.Run(new List<PersonFiles>(files));
        }

        public PersonFiles[] SplitExcel(string docName, byte[] docFile, byte[] xmlFile)
        {
            ISplit run = new ExcelSplit(docName);
            MemoryStream doc = new MemoryStream(docFile);
            MemoryStream xml = new MemoryStream(xmlFile);
            run.OpenAndSearchDocument(doc, xml);

            return run.SaveSplitDocument(doc).ToArray();
        }

        public PersonFiles[] SplitPresentation(string docName, byte[] docFile, byte[] xmlFile)
        {
            ISplit run = new PresentationSplit(docName);
            MemoryStream doc = new MemoryStream(docFile);
            MemoryStream xml = new MemoryStream(xmlFile);
            run.OpenAndSearchDocument(doc, xml);

            return run.SaveSplitDocument(doc).ToArray();
        }

        public PersonFiles[] SplitWord(string docName, byte[] docxFile, byte[] xmlFile)
        {
            ISplit run = new WordSplit(docName);
            MemoryStream doc = new MemoryStream(docxFile);
            MemoryStream xml = new MemoryStream(xmlFile);
            run.OpenAndSearchDocument(doc, xml);

            return run.SaveSplitDocument(doc).ToArray();
        }
    }
}
