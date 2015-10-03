using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Reflection;

using DocumentFormat.OpenXml.Packaging;
using Wordproc = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using SplitDescriptionObjects;

namespace DocumentSplitEngine
{
	public class OpenXMDocumentPart
	{
		public IList<OpenXmlElement> CompositeElements { get; set; }
		public string PartOwner { get; set; }
		public Guid Guid { get; private set; }

		public OpenXMDocumentPart()
		{
			this.Guid = Guid.NewGuid();
			CompositeElements = new List<OpenXmlElement>();
		}
	}

	public abstract class MarkerMapper
	{
		protected Split Xml { get; set; }
		protected string[] SubdividedParagraphs { get; set; }
	}

	public interface IMarkerMapper
	{
		IList<OpenXMDocumentPart> Run();
    }

	public class MarkerDocumentMapper : MarkerMapper, IMarkerMapper
	{	
		SplitDocument SplitDocumentObj { get; set; }
		Wordproc.Body DocumentBody { get; set; }
		
		public MarkerDocumentMapper(string documentName, Split xml, Wordproc.Body body)
		{
			Xml = xml;
			SplitDocumentObj = (SplitDocument)Xml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, documentName)).SingleOrDefault();
			DocumentBody = body;
			SubdividedParagraphs = new string[body.ChildElements.Count];
		}

		public IUniversalDocumentMarker GetUniversalDocumentMarker()
		{
			return new UniversalDocumentMarker(DocumentBody);
		}

		public IList<OpenXMDocumentPart> Run()
		{
			IList<OpenXMDocumentPart> documentElements = new List<OpenXMDocumentPart>();
			if (SplitDocumentObj != null)
			{
				foreach (Person person in SplitDocumentObj.Person)
				{
					if (person.UniversalMarker != null)
					{
						foreach (PersonUniversalMarker marker in person.UniversalMarker)
						{
							IList<int> result = GetUniversalDocumentMarker().GetCrossedElements(marker.ElementId, marker.SelectionLastelementId);
							foreach (int index in result)
							{
								if (string.IsNullOrEmpty(SubdividedParagraphs[index]))
								{
									SubdividedParagraphs[index] = person.Email;
								}
								else
									throw new ElementToPersonPairException();
							}
						}
					}

					if (person.TextMarker != null)
					{
					}

					if (person.PictureMarker != null)
					{
					}

					if (person.TableMarker != null)
					{
					}
				}

				string email = string.Empty;
				OpenXMDocumentPart part = new OpenXMDocumentPart();				
				for (int index = 0; index < DocumentBody.ChildElements.Count; index++)
				{
					if (SubdividedParagraphs[index] != email)
					{
						part = new OpenXMDocumentPart();
						part.CompositeElements.Add(DocumentBody.ChildElements[index]);
						email = SubdividedParagraphs[index];
						if (string.IsNullOrEmpty(email))
							part.PartOwner = "undefined";
						else
							part.PartOwner = email;
						
						documentElements.Add(part);
					}
					else
						part.CompositeElements.Add(DocumentBody.ChildElements[index]);
				}
			}

			return documentElements;
		}
	}

	public interface ILocalSplit
	{
		void OpenAndSearchWordDocument(string filePath, string xmlSplitDefinitionFilePath);		
		void SaveSplitDocument(string filePath);		
	}

	public interface ISplit
	{
		List<PersonFiles> SaveSplitDocument(Stream document);
		void OpenAndSearchWordDocument(Stream docxFile, Stream xmlFile);
	}

	public interface IMergeXml
	{
		void CreateMergeXml(string path);
		byte[] CreateMergeXml();
    }

	public class MergeXml : IMergeXml
	{
		protected IList<OpenXMDocumentPart> DocumentElements;
		protected string DocumentName { get; set; }

		public void CreateMergeXml(string path)
		{
			Merge mergeXml = new Merge();
			mergeXml.Items = new MergeDocument[1];
			mergeXml.Items[0] = new MergeDocument();
			mergeXml.Items[0].Name = DocumentName;
			mergeXml.Items[0].Part = new MergeDocumentPart[DocumentElements.Count];
			for (int index = 0; index < DocumentElements.Count; index++)
			{
				mergeXml.Items[0].Part[index] = new MergeDocumentPart();
				mergeXml.Items[0].Part[index].Name = DocumentElements[index].PartOwner;
				mergeXml.Items[0].Part[index].Id = DocumentElements[index].Guid.ToString();
			}

			using (FileStream fileStream = new FileStream(path + "mergeXmlDefinition" + ".xml",
							System.IO.FileMode.CreateNew))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Merge));
				serializer.Serialize(fileStream, mergeXml);
			}
		}

		public byte[] CreateMergeXml()
		{
			Merge mergeXml = new Merge();
			mergeXml.Items = new MergeDocument[1];
			mergeXml.Items[0] = new MergeDocument();
			mergeXml.Items[0].Name = DocumentName;
			mergeXml.Items[0].Part = new MergeDocumentPart[DocumentElements.Count];
			for (int index = 0; index < DocumentElements.Count; index++)
			{
				mergeXml.Items[0].Part[index] = new MergeDocumentPart();
				mergeXml.Items[0].Part[index].Name = DocumentElements[index].PartOwner;
				mergeXml.Items[0].Part[index].Id = DocumentElements[index].Guid.ToString();
			}

			using (MemoryStream mergeStream = new MemoryStream())
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Merge));
				serializer.Serialize(mergeStream, mergeXml);

				return mergeStream.ToArray();
			}
		}

		public static byte[] ReadFully(Stream stream)
		{
			long originalPosition = 0;

			if (stream.CanSeek)
			{
				originalPosition = stream.Position;
				stream.Position = 0;
			}

			try
			{
				byte[] readBuffer = new byte[4096];

				int totalBytesRead = 0;
				int bytesRead;

				while ((bytesRead = stream.Read(readBuffer, totalBytesRead, readBuffer.Length - totalBytesRead)) > 0)
				{
					totalBytesRead += bytesRead;

					if (totalBytesRead == readBuffer.Length)
					{
						int nextByte = stream.ReadByte();
						if (nextByte != -1)
						{
							byte[] temp = new byte[readBuffer.Length * 2];
							Buffer.BlockCopy(readBuffer, 0, temp, 0, readBuffer.Length);
							Buffer.SetByte(temp, totalBytesRead, (byte)nextByte);
							readBuffer = temp;
							totalBytesRead++;
						}
					}
				}

				byte[] buffer = readBuffer;
				if (readBuffer.Length != totalBytesRead)
				{
					buffer = new byte[totalBytesRead];
					Buffer.BlockCopy(readBuffer, 0, buffer, 0, totalBytesRead);
				}
				return buffer;
			}
			finally
			{
				if (stream.CanSeek)
				{
					stream.Position = originalPosition;
				}
			}
		}
	}

	public class DocumentSplit : MergeXml, ISplit, ILocalSplit
    {
		public DocumentSplit(string docName)
		{
			DocumentName = docName;
		}

		public void OpenAndSearchWordDocument(Stream docxFile, Stream xmlFile)
		{
			XmlSerializer serializer = new XmlSerializer(typeof(Split));
			Split splitXml = (Split)serializer.Deserialize(xmlFile);
			using (WordprocessingDocument wordDoc =
				WordprocessingDocument.Open(docxFile, true))
			{
				Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
				IMarkerMapper mapping = new MarkerDocumentMapper(DocumentName, splitXml, body);
				DocumentElements = mapping.Run();
			}
		}

		public void OpenAndSearchWordDocument(string docxFilePath, string xmlFilePath)
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
			MarkerDocumentMapper mapping = new MarkerDocumentMapper(DocumentName, splitXml, body);
			DocumentElements = mapping.Run();

			// Close the handle explicitly.
			wordprocessingDocument.Close();
		}

		public void SaveSplitDocument(string docxFilePath)
		{
			DirectoryInfo initDi;
			string appPath = Path.GetDirectoryName(Assembly.GetAssembly(typeof(DocumentSplit)).Location);
			if (!Directory.Exists(appPath + @"\Files"))
				initDi = Directory.CreateDirectory(appPath + @"\Files");

			byte[] byteArray = File.ReadAllBytes(docxFilePath);
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					foreach (OpenXMDocumentPart element in DocumentElements)
					{						
						wordDoc.MainDocumentPart.Document.Body = new Wordproc.Body();
						Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
						foreach (OpenXmlElement compo in element.CompositeElements)
							body.Append(compo.CloneNode(true));

						wordDoc.MainDocumentPart.Document.Save();

						string directoryPath = appPath + @"\Files" + @"\" + element.PartOwner;
						DirectoryInfo currentDi;
						if (!Directory.Exists(directoryPath))
						{
							currentDi = Directory.CreateDirectory(directoryPath);
						}

						using (FileStream fileStream = new FileStream(directoryPath + @"\" + element.Guid.ToString() + ".docx",
							System.IO.FileMode.CreateNew))
						{
							mem.WriteTo(fileStream);
						}
					}
				}
				// At this point, the memory stream contains the modified document.
				// We could write it back to a SharePoint document library or serve
				// it from a web server.			
			}

			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					wordDoc.MainDocumentPart.Document.Body = new Wordproc.Body();
					wordDoc.MainDocumentPart.Document.Save();

					using (FileStream fileStream = new FileStream(appPath + @"\Files" + @"\template" + ".docx",
						System.IO.FileMode.CreateNew))
					{
						mem.WriteTo(fileStream);
					}
				}
				// At this point, the memory stream contains the modified document.
				// We could write it back to a SharePoint document library or serve
				// it from a web server.			
			}

			CreateMergeXml(appPath + @"\Files" + @"\");
		}

		List<PersonFiles> ISplit.SaveSplitDocument(Stream document)
		{
			List<PersonFiles> resultList = new List<PersonFiles>();

			byte[] byteArray = ReadFully(document);
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					foreach (OpenXMDocumentPart element in DocumentElements)
					{
						wordDoc.MainDocumentPart.Document.Body = new Wordproc.Body();
						Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
						foreach (OpenXmlElement compo in element.CompositeElements)
							body.Append(compo.CloneNode(true));

						wordDoc.MainDocumentPart.Document.Save();

						var person = new PersonFiles();
						person.Person = element.PartOwner;
						resultList.Add(person);
						person.Name = element.Guid.ToString();
						person.Data = mem.ToArray();
					}
				}
			}
			// At this point, the memory stream contains the modified document.
			// We could write it back to a SharePoint document library or serve
			// it from a web server.			
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument wordDoc =
					WordprocessingDocument.Open(mem, true))
				{
					wordDoc.MainDocumentPart.Document.Body = new Wordproc.Body();
					wordDoc.MainDocumentPart.Document.Save();

					var person = new PersonFiles();
					person.Person = "/";
					resultList.Add(person);
					person.Name = "template.docx";
					person.Data = mem.ToArray();
				}
			}
			// At this point, the memory stream contains the modified document.
			// We could write it back to a SharePoint document library or serve
			// it from a web server.			

			var xmlPerson = new PersonFiles();
			xmlPerson.Person = "/";
			resultList.Add(xmlPerson);
			xmlPerson.Name = "mergeXmlDefinition.xml";
			xmlPerson.Data = CreateMergeXml();

			return resultList;
		}
	}
}
