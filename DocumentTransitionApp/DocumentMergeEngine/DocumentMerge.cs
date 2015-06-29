using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocumentMergeEngine
{
    public class DocumentMerge
    {
		string DocumentPath { get; set; }

		public DocumentMerge(string path)
		{
			DocumentPath = path;
		}

		public void Run()
		{
			string appPath = Path.GetDirectoryName(Assembly.GetAssembly(typeof(DocumentMerge)).Location);
			string xmlFilePath = appPath + @"\Files\" + "mergeXmlDefinition.xml";
			var xml = System.IO.File.ReadAllText(xmlFilePath);
			Merge mergeXml;
			using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Merge));
				mergeXml = (Merge)serializer.Deserialize(stream);
			}

			Body body = new Body();
			MergeDocument documentXml = mergeXml.Items.First();
			foreach (MergeDocumentPart part in documentXml.Part)
			{
				WordprocessingDocument wordprocessingDocument =
					WordprocessingDocument.Open(appPath + @"\Files\" + part.Name + @"\" + part.Id + ".docx", true);

				// Assign a reference to the existing document body.
				foreach (OpenXmlElement element in wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements)
				{
					body.Append(element.CloneNode(true));
				}

				// Close the handle explicitly.
				wordprocessingDocument.Close();
			}

			byte[] byteArray = File.ReadAllBytes(appPath + @"\Files\" + "template.docx");
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					wordDoc.MainDocumentPart.Document.Body = body;
					wordDoc.MainDocumentPart.Document.Save();

					using (FileStream fileStream = new FileStream(DocumentPath,
						System.IO.FileMode.CreateNew))
					{
						mem.WriteTo(fileStream);
					}
				}
			}
		}

		public void CreateDocumentPart()
		{
			// Create a Wordprocessing document. 
			using (WordprocessingDocument myDoc = WordprocessingDocument.Create(DocumentPath, WordprocessingDocumentType.Document))
			{
				// Add a new main document part. 
				MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
				//Create DOM tree for simple document. 
				mainPart.Document = new Document();
				Body body = new Body();
				Paragraph p = new Paragraph();
				Run r = new Run();
				Text t = new Text("Hello World!");
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
