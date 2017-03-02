using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace DocumentMergeEngine
{
    public interface IDocumentMerge
    {
        Merge GetMergeXml(List<PersonFiles> files);
    }

    public abstract class DocumentMerge : IDocumentMerge
    {
        public Merge GetMergeXml(List<PersonFiles> files)
        {
            var xml = files.Where(p => p.Person == "/" && p.Name == "mergeXmlDefinition.xml").Select(d => d.Data).FirstOrDefault();
            Merge mergeXml;
            using (MemoryStream stream = new MemoryStream(xml))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Merge));
                mergeXml = (Merge)serializer.Deserialize(stream);
            }

            return mergeXml;
        }
    }
}
