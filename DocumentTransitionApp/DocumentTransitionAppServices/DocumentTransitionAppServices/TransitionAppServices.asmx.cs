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
		public PersonFiles[] SplitDocument(string docName, byte[] docxFile, byte[] xmlFile)
		{
			ISplit run = new DocumentSplit(docName);
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
        public byte[] GenerateSplitDocument(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplitXml split = new DocumentSplit(docName);

            return split.CreateSplitXml(parts);
        }

		[WebMethod]
		public byte[] MergeDocument(PersonFiles[] files)
		{
			IMerge merge = new DocumentMerge();

			return merge.Run(new List<PersonFiles>(files));
		}

		[WebMethod]
		public List<PartsSelectionTreeElement> GetDocumentParts(string docName, byte[] documentFile)
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
        public ServiceResponse GetDocumentPartsFromXml(string docName, byte[] documentFile, byte[] splitFile)
        {
            ISplitXml split = new DocumentSplit(Path.GetFileNameWithoutExtension(docName));
            var cleanParts = GetDocumentParts(docName, documentFile);
            
            try
            {
                return new ServiceResponse(split.PartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
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
                return new ServiceResponse(split.PartsFromSplitXml(new MemoryStream(splitFile), cleanParts));
            }
            catch (SplitNameDifferenceExcception ex)
            {
                return new ServiceResponse(ex.Message);
            }
        }
    }	
}