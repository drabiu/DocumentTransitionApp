﻿using System;
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
using DocumentMergeEngine.Interfaces;

namespace DocumentMergeEngine
{
    public class WordMerge : DocumentMerge, ILocalMerge, IMerge
    {
        [Obsolete]
		public void Run(string path)
		{
			string appPath = Path.GetDirectoryName(Assembly.GetAssembly(typeof(WordMerge)).Location);
			string xmlFilePath = appPath + @"\Files\" + "mergeXmlDefinition.xml";
			var xml = File.ReadAllText(xmlFilePath);
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
                        FileMode.CreateNew))
					{
						mem.WriteTo(fileStream);
					}
				}
			}
		}

		public byte[] Run(List<PersonFiles> files)
		{
            var mergeXml = GetMergeXml(files);

			Body body = new Body();
			MergeDocument documentXml = mergeXml.Items.First();
			foreach (MergeDocumentPart part in documentXml.Part)
			{
				byte[] byteArray = files.Where(p => p.Person == part.Name && p.Name == part.Id).Select(d => d.Data).FirstOrDefault();
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

			byte[] template = files.Where(p => p.Person == "/" && p.Name == "template.docx").Select(d => d.Data).FirstOrDefault();
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
