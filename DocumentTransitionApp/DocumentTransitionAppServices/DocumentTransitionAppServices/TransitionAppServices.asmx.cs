using System.Collections.Generic;
using System.IO;
using System.Web.Services;

using DocumentSplitEngine;
using SplitDescriptionObjects;
using DocumentMergeEngine;
using DocumentEditPartsEngine;

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
	public class Service1 : System.Web.Services.WebService
	{
		[WebMethod]
		public PersonFiles[] SplitDocument(string docName, byte[] docxFile, byte[] xmlFile)
		{
			ISplit run = new DocumentSplit(docName);
			MemoryStream doc = new MemoryStream(docxFile);
			MemoryStream xml = new MemoryStream(xmlFile);
			run.OpenAndSearchWordDocument(doc, xml);

			return run.SaveSplitDocument(doc).ToArray();
		}

        [WebMethod]
        public byte[] GenerateSplitDocument(string docName, PartsSelectionTreeElement[] parts)
        {
            ISplit split = new DocumentSplit(docName);

            return split.CreateSplitXml(parts);
        }

		[WebMethod]
		public byte[] MergeDocument(string docName, PersonFiles[] files)
		{
			IMerge merge = new DocumentMerge();

			return merge.Run(new List<PersonFiles>(files));
		}

		[WebMethod]
		public List<PartsSelectionTreeElement> GetParts(string docName, byte[] documentFile)
		{
			IDocumentParts parts = DocumentPartsBuilder.Build(Path.GetExtension(docName));

			return parts.Get(new MemoryStream(documentFile));
		}
	}	
}