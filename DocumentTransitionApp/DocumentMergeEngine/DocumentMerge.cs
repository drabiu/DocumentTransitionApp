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
using SplitDescriptionObjects;

namespace DocumentMergeEngine
{
	public interface ILocalMerge
	{
		void Run(string path);
	}

	public interface IMerge
	{
		byte[] Run(List<PersonFiles> files);
	}

    public class DocumentMerge : ILocalMerge, IMerge
    {
		public void Run(string path)
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

					using (FileStream fileStream = new FileStream(path,
						System.IO.FileMode.CreateNew))
					{
						mem.WriteTo(fileStream);
					}
				}
			}
		}

		public byte[] Run(List<PersonFiles> files)
		{
			var xml = files.FirstOrDefault(p => p.Person == "/").Files.FirstOrDefault(f => f.Name == "mergeXmlDefinition.xml").Data;
			Merge mergeXml;
			using (MemoryStream stream = new MemoryStream(xml))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Merge));
				mergeXml = (Merge)serializer.Deserialize(stream);
			}

			Body body = new Body();
			MergeDocument documentXml = mergeXml.Items.First();
			foreach (MergeDocumentPart part in documentXml.Part)
			{
				byte[] byteArray = files.FirstOrDefault(p => p.Person == part.Name).Files.FirstOrDefault(f => f.Name == part.Id + ".docx").Data;
				using (MemoryStream mem = new MemoryStream())
				{
					mem.Write(byteArray, 0, (int)byteArray.Length);
					WordprocessingDocument wordprocessingDocument =
						WordprocessingDocument.Open(mem, true);

					// Assign a reference to the existing document body.
					foreach (OpenXmlElement element in wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements)
					{
						body.Append(element.CloneNode(true));
					}

					// Close the handle explicitly.
					wordprocessingDocument.Close();
				}
			}

			byte[] template = files.FirstOrDefault(p => p.Person == "/").Files.FirstOrDefault(f => f.Name == "template.docx").Data;
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(template, 0, (int)template.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					wordDoc.MainDocumentPart.Document.Body = body;
					wordDoc.MainDocumentPart.Document.Save();

					return mem.ToArray();
				}
			}
		}
	}
}
