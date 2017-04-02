using DocumentFormat.OpenXml.Packaging;
using DocumentMergeEngine.Interfaces;
using OpenXmlPowerTools;
using OpenXMLTools;
using OpenXMLTools.Interfaces;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocumentMergeEngine
{
    public class PresentationMerge : DocumentMerge, IMerge
    {
        public byte[] Run(List<PersonFiles> files)
        {
            var mergeXml = GetMergeXml(files);
            MergeDocument documentXml = mergeXml.Items.First();
            IPresentationTools presentationTools = new PresentationTools();

            byte[] emptyTemplate = files.Where(p => p.Person == "/" && p.Name == "template.pptx").Select(d => d.Data).FirstOrDefault();
            using (MemoryStream emptyDocInMemoryStream = new MemoryStream(emptyTemplate, 0, emptyTemplate.Length, true, true))
            {
                OpenXmlPowerToolsDocument emptyDocPowerTools = new OpenXmlPowerToolsDocument(string.Empty, emptyDocInMemoryStream);
                using (OpenXmlMemoryStreamDocument streamEmptyDoc = new OpenXmlMemoryStreamDocument(emptyDocPowerTools))
                {
                    PresentationDocument emptyPresentation = streamEmptyDoc.GetPresentationDocument();
                    foreach (MergeDocumentPart part in documentXml.Part)
                    {
                        byte[] byteArray = files.Where(p => p.Person == part.Name.Trim() && p.Name == part.Id).Select(d => d.Data).FirstOrDefault();
                        using (MemoryStream partDocInMemoryStream = new MemoryStream(byteArray, 0, byteArray.Length, true, true))
                        {
                            OpenXmlPowerToolsDocument partDocPowerTools = new OpenXmlPowerToolsDocument(string.Empty, partDocInMemoryStream);
                            using (OpenXmlMemoryStreamDocument streamDividedDoc = new OpenXmlMemoryStreamDocument(partDocPowerTools))
                            {
                                PresentationDocument templatePresentation = streamDividedDoc.GetPresentationDocument();
                                presentationTools.InsertSlidesFromTemplate(emptyPresentation, templatePresentation);
                            }
                        }
                    }

                    return streamEmptyDoc.GetModifiedDocument().DocumentByteArray;
                }
            }
        }
    }
}
