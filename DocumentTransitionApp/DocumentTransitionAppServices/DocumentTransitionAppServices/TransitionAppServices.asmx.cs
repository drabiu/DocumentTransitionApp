using DocumentEditPartsEngine;
using DocumentEditPartsEngine.Interfaces;
using DocumentMergeEngine;
using DocumentMergeEngine.Interfaces;
using DocumentSplitEngine;
using DocumentSplitEngine.Interfaces;
using DocumentTransitionAppServices.Interfaces;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;
using System.Web.Services;
using System;

namespace DocumentTransitionAppServices
{
    /// <summary>
    /// Summary description for Service1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
	[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
	[System.ComponentModel.ToolboxItem(false)]
	// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
	// [System.Web.Script.Services.ScriptService]
	public class Service1 : WebService, ITransitionAppService
    {
		[WebMethod]
		public PersonFiles[] SplitWord(string docName, byte[] docxFile, byte[] xmlFile)
		{
			ISplit run = new WordSplit(docName);
			MemoryStream doc = new MemoryStream(docxFile);
			MemoryStream xml = new MemoryStream(xmlFile);
			run.OpenAndSearchDocument(doc, xml);

			return run.SaveSplitDocument(doc).ToArray();
		}

        [WebMethod]
        public PersonFiles[] SplitPresentation(string docName, byte[] docFile, byte[] xmlFile)
        {
            ISplit run = new PresentationSplit(docName);
            MemoryStream doc = new MemoryStream(docFile);
            MemoryStream xml = new MemoryStream(xmlFile);
            run.OpenAndSearchDocument(doc, xml);

            return run.SaveSplitDocument(doc).ToArray();
        }

        [WebMethod]
        public byte[] GenerateSplitWord(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplitXml split = new WordSplit(docName);

            return split.CreateSplitXml(parts);
        }

        [WebMethod]
        public byte[] GenerateSplitPresentation(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplitXml split = new PresentationSplit(docName);

            return split.CreateSplitXml(parts);
        }

		[WebMethod]
		public List<PartsSelectionTreeElement> GetWordParts(string docName, byte[] documentFile)
		{
            IDocumentParts parts = new WordDocumentParts();

			return parts.Get(new MemoryStream(documentFile));
		}

        [WebMethod]
        public List<PartsSelectionTreeElement> GetPresentationParts(string preName, byte[] presentationFile)
        {
            IPresentationParts parts = new PresentationDocumentParts();

            return parts.GetSlides(new MemoryStream(presentationFile));
        }

        [WebMethod]
        public ServiceResponse GetWordPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            ISplitXml split = new WordSplit(Path.GetFileNameWithoutExtension(docName));
            var cleanParts = GetWordParts(docName, documentFile);
            
            try
            {
                return new ServiceResponse(split.SelectPartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
            }
            catch (SplitNameDifferenceExcception ex)
            {
                return new ServiceResponse(ex.Message);
            }
        }

        [WebMethod]
        public ServiceResponse GetPresentationPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            ISplitXml split = new PresentationSplit(Path.GetFileNameWithoutExtension(docName));
            var cleanParts = GetPresentationParts(docName, documentFile);

            try
            {
                return new ServiceResponse(split.SelectPartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
            }
            catch (SplitNameDifferenceExcception ex)
            {
                return new ServiceResponse(ex.Message);
            }
        }

        [WebMethod]
        public PersonFiles[] SplitExcel(string docName, byte[] docFile, byte[] xmlFile)
        {
            throw new NotImplementedException();
        }

        [WebMethod]
        public byte[] GenerateSplitExcel(string docName, PartsSelectionTreeElement[] parts)
        {
            throw new NotImplementedException();
        }

        [WebMethod]
        public List<PartsSelectionTreeElement> GetExcelParts(string excName, byte[] excelFile)
        {
            throw new NotImplementedException();
        }

        [WebMethod]
        public ServiceResponse GetExcelPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            throw new NotImplementedException();
        }

        [WebMethod]
        public byte[] MergeWord(PersonFiles[] files)
        {
            IMerge merge = new WordMerge();

            return merge.Run(new List<PersonFiles>(files));
        }

        [WebMethod]
        public byte[] MergePresentation(PersonFiles[] files)
        {
            IMerge merge = new PresentationMerge();

            return merge.Run(new List<PersonFiles>(files));
        }

        [WebMethod]
        public byte[] MergeExcel(PersonFiles[] files)
        {
            throw new NotImplementedException();
        }
    }	
}