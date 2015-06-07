using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;

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

			// Open a WordprocessingDocument for editing using the filepath.
			//WordprocessingDocument wordprocessingDocument =
			//	WordprocessingDocument.Open(docxFilePath, true);

			//// Assign a reference to the existing document body.
			//Wordproc.Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
			//MarkerMapper mapping = new MarkerMapper(DocumentName, splitXml, body);
			//DocumentElements = mapping.Run();

			//// Close the handle explicitly.
			//wordprocessingDocument.Close();
		}
    }
}
