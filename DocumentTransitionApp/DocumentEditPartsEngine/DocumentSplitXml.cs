using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentEditPartsEngine
{
    public interface IDocumentSplitXml
    {
        byte[] CreateSplitXml(string docName, IList<PartsSelectionTreeElement> parts);
    }

    public class WordSplitXml : IDocumentSplitXml
    {
        public byte[] CreateSplitXml(string docName, IList<PartsSelectionTreeElement> parts)
        {
            Split splitXml = new Split();
            splitXml.Items = new SplitDocument[1];
            splitXml.Items[0] = new SplitDocument();
            (splitXml.Items[0] as SplitDocument).Name = docName;
            //mergeXml.Items[0].Part = new MergeDocumentPart[DocumentElements.Count];
            //for (int index = 0; index < DocumentElements.Count; index++)
            //{
            //    mergeXml.Items[0].Part[index] = new MergeDocumentPart();
            //    mergeXml.Items[0].Part[index].Name = DocumentElements[index].PartOwner;
            //    mergeXml.Items[0].Part[index].Id = DocumentElements[index].Guid.ToString();
            //}

            //using (MemoryStream mergeStream = new MemoryStream())
            //{
            //    XmlSerializer serializer = new XmlSerializer(typeof(Merge));
            //    serializer.Serialize(mergeStream, mergeXml);

            //    return mergeStream.ToArray();
            //}
        }
    }
}
