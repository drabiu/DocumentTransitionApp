using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentSplitEngine.Document;
using DocumentSplitEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class MarkerDocumentMapperTest
    {
        IMarkerMapper<OpenXmlElement> MarkerDocumentMapper;
        WordprocessingDocument WordDemoDoc;

        [TestInitialize]
        public void Init()
        {
            byte[] sampleXmlBinary = File.ReadAllBytes(@"../../../Files/split_demo.docx_20170227215840894.xml");

            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            Split splitXml = (Split)serializer.Deserialize(new MemoryStream(sampleXmlBinary));

            WordDemoDoc = WordprocessingDocument.Open(@"../../../Files/demo.docx", false);

            MarkerDocumentMapper = new MarkerDocumentMapper("demo", splitXml, WordDemoDoc.MainDocumentPart.Document.Body);
        }

        [TestMethod]
        public void RunShouldReturn6Parts()
        {
            var documentPartList = MarkerDocumentMapper.Run();

            Assert.AreEqual(7, documentPartList.Count);
        }

        [TestMethod]
        public void RunShouldReturn1ElementForPartByOwner()
        {
            var documentPartList = MarkerDocumentMapper.Run();
            var ownerCompositeElements = documentPartList.SingleOrDefault(p => p.PartOwner == "test").CompositeElements;

            Assert.AreEqual(3, ownerCompositeElements.Count);
        }
        [TestMethod]
        public void RunShouldReturn1ElementForEachPartByOwner()
        {
            var documentPartList = MarkerDocumentMapper.Run();
            var ownerCompositeElements = documentPartList.Where(p => p.PartOwner == "test2").Select(o => o.CompositeElements);

            Assert.AreEqual(2, ownerCompositeElements.Count());
            Assert.AreEqual(1, ownerCompositeElements.ElementAt(0).Count);
            Assert.AreEqual(1, ownerCompositeElements.ElementAt(1).Count);
        }

        [TestCleanup]
        public void Finish()
        {
            WordDemoDoc.Close();
        }
    }
}
