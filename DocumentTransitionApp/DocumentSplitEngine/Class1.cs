using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

using DocumentFormat.OpenXml.Packaging;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using UnmarshallingSplitXml;

namespace DocumentSplitEngine
{
	public class OpenXMDocumentPart
	{
		IList<OpenXmlCompositeElement> CompositeElements { get; set; }
	}

    public class Class1
    {
		IList<OpenXMDocumentPart> DocumentElements;
		bool[] SubdividedParagraphs;

		public static void OpenAndSearchWordDocument(string docxFilePath, string xmlFilePath)
		{
			//split XML Read
			var xml = System.IO.File.ReadAllText(xmlFilePath);
			Split splitXml;
			using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Split));
				splitXml = (Split)serializer.Deserialize(stream);
			}

			// Open a WordprocessingDocument for editing using the filepath.
			WordprocessingDocument wordprocessingDocument =
				WordprocessingDocument.Open(docxFilePath, true);

			// Assign a reference to the existing document body.
			Wordproc.Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
			for (int index = 0; index < body.ChildElements.Count; index++)
			{
			}

			// Close the handle explicitly.
			wordprocessingDocument.Close();
		}

		public static void CreateDocumentPart(OpenXMDocumentPart documentPart)
		{
			// Create a Wordprocessing document. 
			using (WordprocessingDocument myDoc = WordprocessingDocument.Create(AppDomain.CurrentDomain.BaseDirectory + Guid.NewGuid().ToString(), WordprocessingDocumentType.Document))
			{
				// Add a new main document part. 
				MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
				//Create DOM tree for simple document. 
				mainPart.Document = new Wordproc.Document();
				Wordproc.Body body = new Wordproc.Body();
				Wordproc.Paragraph p = new Wordproc.Paragraph();
				Wordproc.Run r = new Wordproc.Run();
				Wordproc.Text t = new Wordproc.Text("Hello World!");
				//Append elements appropriately. 
				r.Append(t);
				p.Append(r);
				body.Append(p);
				mainPart.Document.Append(body);
				// Save changes to the main document part. 
				mainPart.Document.Save();
			}
		}
    }
}
