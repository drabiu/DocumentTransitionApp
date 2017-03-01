using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentSplitEngine.Interfaces;
using DocumentSplitEngine.Presentation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using UnmarshallingSplitXml;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class MarkerPresentationMapperTest
    {
        IMarkerMapper<SlideId> MarkerPresentationMapper;
        PresentationDocument PreSampleDoc;

        [TestInitialize]
        public void Init()
        {
            byte[] sampleXmlBinary = File.ReadAllBytes(@"../../../Files/split_przykladowa-prezentacja.pptx_20170227215707619.xml");

            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            Split splitXml = (Split)serializer.Deserialize(new MemoryStream(sampleXmlBinary));

            PreSampleDoc = PresentationDocument.Open(@"../../../Files/przykladowa-prezentacja.pptx", false);

            MarkerPresentationMapper = new MarkerPresentationMapper("przykladowa-prezentacja", splitXml, PreSampleDoc.PresentationPart);
        }

        [TestMethod]
        public void RunShouldReturn6Parts()
        {
            var documentPartList = MarkerPresentationMapper.Run();

            Assert.AreEqual(6, documentPartList.Count);
        }

        [TestMethod]
        public void RunShouldReturn3ElementForPartByOwner()
        {
            var documentPartList = MarkerPresentationMapper.Run();
            var ownerCompositeElements = documentPartList.SingleOrDefault(p => p.PartOwner == "test").CompositeElements;

            Assert.AreEqual(3, ownerCompositeElements.Count);
        }
        [TestMethod]
        public void RunShouldReturn1ElementForEachPartByOwner()
        {
            var documentPartList = MarkerPresentationMapper.Run();
            var ownerCompositeElements = documentPartList.Where(p => p.PartOwner == "test2").Select(o => o.CompositeElements);

            Assert.AreEqual(2, ownerCompositeElements.Count());
            Assert.AreEqual(1, ownerCompositeElements.ElementAt(0).Count);
            Assert.AreEqual(1, ownerCompositeElements.ElementAt(1).Count);
        }

        [TestCleanup]
        public void Finish()
        {
            PreSampleDoc.Close();
        }
    }
}
