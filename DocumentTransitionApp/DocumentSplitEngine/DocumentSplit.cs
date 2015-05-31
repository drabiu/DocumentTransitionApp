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

    public class DocumentSplit
    {
		IList<OpenXMDocumentPart> DocumentElements;

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
			MarkerMapper mapping = new MarkerMapper(Path.GetFileNameWithoutExtension(docxFilePath), splitXml, body);
			DocumentElements = mapping.Run();

			// Close the handle explicitly.
			wordprocessingDocument.Close();
		}

		private void CreateMergeXml()
		{
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
						Wordproc.Body body = wordDoc.MainDocumentPart.Document.Body;
						body = new Wordproc.Body();
						foreach (OpenXmlElement compo in element.CompositeElements)
							body.Append(compo.CloneNode(true));

						string directoryPath = appPath + @"\Files" + @"\" + element.PartOwner;
						DirectoryInfo currentDi;
						if (!Directory.Exists(directoryPath))
						{
							currentDi = Directory.CreateDirectory(directoryPath);
						}

						using (FileStream fileStream = new FileStream(directoryPath + @"\" + element.Guid.ToString() + ".docx",
							System.IO.FileMode.CreateNew))
						{
							mem.CopyTo(fileStream);
						}
					}
				}
				// At this point, the memory stream contains the modified document.
				// We could write it back to a SharePoint document library or serve
				// it from a web server.

				// In this example, we serialize back to the file system to verify
				// that the code worked properly.			
			}
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
