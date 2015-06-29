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

using UnmarshallingSplitXml;
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

	public class MarkerMapper
	{
		Split Xml { get; set; }
		SplitDocument SplitDocumentObj { get; set; }
		Wordproc.Body DocumentBody { get; set; }
		string[] SubdividedParagraphs { get; set; }

		public MarkerMapper(string documentName, Split xml, Wordproc.Body body)
		{
			Xml = xml;
			SplitDocumentObj = Xml.Items.Where(it => string.Equals(it.Name, documentName)).SingleOrDefault();
			DocumentBody = body;
			SubdividedParagraphs = new string[body.ChildElements.Count];
		}

		public UniversalDocumentMarker GetEquivalentMarker(SplitDocumentPersonUniversalMarker marker)
		{
			return new UniversalDocumentMarker(DocumentBody);
		}

		public IList<OpenXMDocumentPart> Run()
		{
			IList<OpenXMDocumentPart> documentElements = new List<OpenXMDocumentPart>();
			if (SplitDocumentObj != null)
			{
				foreach (SplitDocumentPerson person in SplitDocumentObj.Person)
				{
					if (person.UniversalMarker != null)
					{
						foreach (SplitDocumentPersonUniversalMarker marker in person.UniversalMarker)
						{
							IList<int> result = GetEquivalentMarker(marker).GetCrossedElements(marker.ElementId, marker.SelectionLastelementId);
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

					part.CompositeElements.Add(DocumentBody.ChildElements[index]);
				}
			}

			return documentElements;
		}
	}

	interface ISplit
	{
		void OpenAndSearchWordDocument(string filePath, string xmlSplitDefinitionFilePath);
		void SaveSplitDocument(string filePath);
	}

	public class MergeXml
	{
		protected IList<OpenXMDocumentPart> DocumentElements;
		protected string DocumentName { get; set; }

		protected void CreateMergeXml(string path)
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
	}

    public class DocumentSplit : MergeXml, ISplit
    {
		public DocumentSplit(string docxFilePath)
		{
			DocumentName = Path.GetFileNameWithoutExtension(docxFilePath);

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
			MarkerMapper mapping = new MarkerMapper(DocumentName, splitXml, body);
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
	}

	public class ExcelSplit : MergeXml, ISplit
	{

		public void OpenAndSearchWordDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
			throw new NotImplementedException();
		}

		public void SaveSplitDocument(string filePath)
		{
			throw new NotImplementedException();
		}
	}

	public class PresentationSplit : MergeXml, ISplit
	{
		public void OpenAndSearchWordDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
			throw new NotImplementedException();
		}

		public void SaveSplitDocument(string filePath)
		{
			throw new NotImplementedException();
		}
	}
}
