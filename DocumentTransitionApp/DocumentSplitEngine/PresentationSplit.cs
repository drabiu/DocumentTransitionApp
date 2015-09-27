using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Presentproc = DocumentFormat.OpenXml.Presentation;

using SplitDescriptionObjects;

namespace DocumentSplitEngine
{
	public class MarkerPresentationMapper : MarkerMapper, IMarkerMapper
	{
		SplitPresentation SplitPresentationObj { get; set; }
		Presentproc.Presentation Presentation;

		public MarkerPresentationMapper(string documentName, Split xml, Presentproc.Presentation presentation)
		{
			Xml = xml;
			SplitPresentationObj = (SplitPresentation)Xml.Items.Where(it => it is SplitPresentation && string.Equals(((SplitPresentation)it).Name, documentName)).SingleOrDefault();
			Presentation = presentation;
			SubdividedParagraphs = new string[presentation.ChildElements.Count];
		}

		public IList<OpenXMDocumentPart> Run()
		{
			throw new NotImplementedException();
		}
	}

	public class PresentationSplit : MergeXml, ISplit, ILocalSplit
	{
		public void OpenAndSearchWordDocument(string filePath, string xmlSplitDefinitionFilePath)
		{
			//split XML Read
			var xml = System.IO.File.ReadAllText(xmlSplitDefinitionFilePath);
			Split splitXml;
			using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Split));
				splitXml = (Split)serializer.Deserialize(stream);
			}

			// Open a WordprocessingDocument for editing using the filepath.
			PresentationDocument wordprocessingDocument =
				PresentationDocument.Open(filePath, true);

			// Assign a reference to the existing document body.
			Presentproc.Presentation presentation = wordprocessingDocument.PresentationPart.Presentation;
			IMarkerMapper mapping = new MarkerPresentationMapper(DocumentName, splitXml, presentation);
			DocumentElements = mapping.Run();

			// Close the handle explicitly.
			wordprocessingDocument.Close();
		}

		public void SaveSplitDocument(string filePath)
		{
			throw new NotImplementedException();
		}

		public void OpenAndSearchWordDocument(Stream docxFile, Stream xmlFile)
		{
			throw new NotImplementedException();
		}

		List<PersonFiles> ISplit.SaveSplitDocument(Stream document)
		{
			throw new NotImplementedException();
		}
	}
}
