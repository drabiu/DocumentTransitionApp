using DocumentEditPartsEngine;
using DocumentSplitEngine.Data_Structures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace DocumentSplitEngine
{
    public interface IMergeXml
    {
        void CreateMergeXml(string path);
        [Obsolete]
        byte[] CreateMergeXml();
    }
      
    public abstract class MergeXml<Element> : IMergeXml
    {
        internal class NameIndexer
        {
            private Dictionary<string, int> Indexes;

            public NameIndexer(IList<string> nameList)
            {
                Indexes = new Dictionary<string, int>();
                foreach (var name in nameList)
                    Indexes.Add(name, 0);
            }

            public int GetNextIndex(string name)
            {
                return Indexes[name]++;
            }
        }

        protected IList<OpenXMLDocumentPart<Element>> DocumentElements;
        protected string DocumentName { get; set; }

        [Obsolete]
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
                            FileMode.CreateNew))
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
    }
}
