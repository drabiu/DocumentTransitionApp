using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SplitDescriptionObjects;
using SplitDescriptionObjects.DocumentMarkers;
using SplitDescriptionObjects.Interfaces;
using SplitDescriptionObjectsTests.Mocks;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace SplitDescriptionObjectsTests
{
    [TestClass]
    public class DocumentMarkersTest
    {
        IUniversalWordMarker UniversalDocNoParagraphIdMarker;
        IUniversalWordMarker UniversalDocParagraphIdMarker;
        ITableWordMarker TableDocNoParagraphIdMarker;
        ITableWordMarker TableDocParagraphIdMarker;
        IListWordMarker ListDocNoParagraphIdMarker;
        IListWordMarker ListDocParagraphIdMarker;
        IPictureWordMarker PictureDocNoParagraphIdMarker;
        IPictureWordMarker PictureDocParagraphIdMarker;
        WordprocessingDocument WordNoParagraphIdDoc;
        WordprocessingDocument WordDemoDoc;
        List<MarkerWordSelector> MarkerWordNoParagraphIdSelectors;
        List<MarkerWordSelector> MarkerWordDemoDocSelectors;

        //ISplitXml SplitXml;
        //byte[] CreateSplitXmlBinary;
        SplitDocument SplitDocumentObj;

        [TestInitialize]
        public void Init()
        {
            WordNoParagraphIdDoc = WordprocessingDocument.Open(@"../../../Files/lep-na-szczury-z-atr-karta-ch.docx", false);
            WordDemoDoc = WordprocessingDocument.Open(@"../../../Files/demo.docx", false);

            byte[] sampleXmlBinary = File.ReadAllBytes(@"../../../Files/split_demo.docx_20170227215840894.xml");

            //SplitXml = new WordSplit("demo");
            //var parts = PartsSelectionTreeElementMock.GetListMock();
            //CreateSplitXmlBinary = SplitXml.CreateSplitXml(parts);
            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            Split splitXml = (Split)serializer.Deserialize(new MemoryStream(sampleXmlBinary));
            SplitDocumentObj = (SplitDocument)splitXml.Items.Where(it => it is SplitDocument && string.Equals(((SplitDocument)it).Name, "demo")).SingleOrDefault();

            //test scenarios when paragraphs have an Id and a paragraph hasn`t got an Id
            MarkerWordNoParagraphIdSelectors = MarkerWordSelector.InitializeSelectorsList(WordNoParagraphIdDoc.MainDocumentPart.Document.Body);
            MarkerWordDemoDocSelectors = MarkerWordSelector.InitializeSelectorsList(WordDemoDoc.MainDocumentPart.Document.Body);
            UniversalDocNoParagraphIdMarker = new UniversalWordMarker(WordNoParagraphIdDoc.MainDocumentPart.Document.Body, MarkerWordNoParagraphIdSelectors);
            UniversalDocParagraphIdMarker = new UniversalWordMarker(WordDemoDoc.MainDocumentPart.Document.Body, MarkerWordDemoDocSelectors);
            TableDocNoParagraphIdMarker = new TableWordMarker(WordNoParagraphIdDoc.MainDocumentPart.Document.Body, MarkerWordNoParagraphIdSelectors);
            TableDocParagraphIdMarker = new TableWordMarker(WordDemoDoc.MainDocumentPart.Document.Body, MarkerWordDemoDocSelectors);
            ListDocNoParagraphIdMarker = new ListWordMarker(WordNoParagraphIdDoc.MainDocumentPart.Document.Body, MarkerWordNoParagraphIdSelectors);
            ListDocParagraphIdMarker = new ListWordMarker(WordDemoDoc.MainDocumentPart.Document.Body, MarkerWordDemoDocSelectors);
            PictureDocNoParagraphIdMarker = new PictureWordMarker(WordNoParagraphIdDoc.MainDocumentPart.Document.Body, MarkerWordNoParagraphIdSelectors);
            PictureDocParagraphIdMarker = new PictureWordMarker(WordDemoDoc.MainDocumentPart.Document.Body, MarkerWordDemoDocSelectors);
        }

        [TestMethod]
        public void UniversalMarkerGetCrossedParagraphElementsShouldReturnOne()
        {
            IList<int> result = UniversalDocNoParagraphIdMarker.GetCrossedParagraphElements(WordDocumentPartAttributes.GetParagraphNoIdFormatter(1), WordDocumentPartAttributes.GetParagraphNoIdFormatter(1));

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(1, result[0]);
        }

        [TestMethod]
        public void UniversalMarkerGetCrossedParagraphElementsShouldReturnThree()
        {
            IList<int> result = UniversalDocParagraphIdMarker.GetCrossedParagraphElements("3CCBE53A", "4424AD34");

            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(17, result[0]);
            Assert.AreEqual(19, result[1]);
            Assert.AreEqual(20, result[2]);
        }

        [TestMethod]
        public void UniversalMarkerGetCrossedParagraphElementsShouldReturnNone()
        {
            IList<int> result = UniversalDocNoParagraphIdMarker.GetCrossedParagraphElements(WordDocumentPartAttributes.GetParagraphNoIdFormatter(4), WordDocumentPartAttributes.GetParagraphNoIdFormatter(1));

            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void UniversalMarkerGetSubdividedPartsShouldReturn3SelectorsForTest()
        {
            var markerWordSelectors = UniversalDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[0]);
            Assert.AreEqual(3, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void UniversalMarkerGetSubdividedPartsShouldReturn2SelectorsForTest2()
        {
            var markerWordSelectors = UniversalDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[1]);
            Assert.AreEqual(2, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void UniversalMarkerGetUniversalMarkersShouldReturn1PersonUniversalMarkerForTest1()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test1");
            var personUniversalMarkers = UniversalWordMarker.GetUniversalMarkers(ownerParts);

            Assert.AreEqual(1, personUniversalMarkers.Count());
            Assert.AreEqual("el1", personUniversalMarkers.First().ElementId);
            Assert.AreEqual("el3", personUniversalMarkers.First().SelectionLastelementId);
        }

        [TestMethod]
        public void UniversalMarkerGetUniversalMarkersShouldReturn2PersonUniversalMarkersForTest2()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test2");
            var personUniversalMarkers = UniversalWordMarker.GetUniversalMarkers(ownerParts);

            Assert.AreEqual(2, personUniversalMarkers.Count());
            Assert.AreEqual("el5", personUniversalMarkers.First().ElementId);
            Assert.AreEqual("el5", personUniversalMarkers.First().SelectionLastelementId);
            Assert.AreEqual("el7", personUniversalMarkers.Skip(1).First().ElementId);
            Assert.AreEqual("el7", personUniversalMarkers.Skip(1).First().SelectionLastelementId);
        }

        [TestMethod]
        public void TableMarkerGetCrossedTableElementsShouldReturn18()
        {
            IList<int> result = TableDocParagraphIdMarker.GetCrossedTableElements(WordDocumentPartAttributes.GetTableIdFormatter(0), WordDocumentPartAttributes.GetTableIdFormatter(0));

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(18, result[0]);
        }

        [TestMethod]
        public void TableMarkerGetCrossedTableElementsShouldReturnNone()
        {
            IList<int> result = TableDocParagraphIdMarker.GetCrossedTableElements(WordDocumentPartAttributes.GetTableIdFormatter(4), WordDocumentPartAttributes.GetTableIdFormatter(0));

            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void TableMarkerGetSubdividedPartsShouldReturn2SelectorsForTest()
        {
            var markerWordSelectors = TableDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[0]);

            Assert.AreEqual(2, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void TableMarkerGetSubdividedPartsShouldReturn1SelectorForTest2()
        {
            var markerWordSelectors = TableDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[1]);

            Assert.AreEqual(1, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void TableMarkerGetTableMarkersShouldReturn2PersonTableMarkersForTest1()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test1");
            var personUniversalMarkers = TableWordMarker.GetTableMarkers(ownerParts);

            Assert.AreEqual(2, personUniversalMarkers.Count());
            Assert.AreEqual("el10", personUniversalMarkers.First().ElementId);
            Assert.AreEqual("el11", personUniversalMarkers.Skip(1).First().ElementId);
        }

        [TestMethod]
        public void TableMarkerGetTableMarkersShouldReturn1PersonTableMarkerForTest2()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test2");
            var personUniversalMarkers = TableWordMarker.GetTableMarkers(ownerParts);

            Assert.AreEqual(1, personUniversalMarkers.Count());
            Assert.AreEqual("el8", personUniversalMarkers.First().ElementId);
        }

        [TestMethod]
        public void ListMarkerGetCrossedListElementsShouldReturn74()
        {
            IList<int> result = ListDocParagraphIdMarker.GetCrossedListElements(WordDocumentPartAttributes.GetParagraphListIdFormatter("3514579C", 2), WordDocumentPartAttributes.GetParagraphListIdFormatter("3514579C", 2));

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(74, result[0]);
        }

        [TestMethod]
        public void ListMarkerGetSubdividedPartsShouldReturn2SelectorsForTest()
        {
            var markerWordSelectors = ListDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[0]);

            Assert.AreEqual(2, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void ListMarkerGetSubdividedPartsShouldReturn2SelectorsForTest2()
        {
            var markerWordSelectors = ListDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[1]);

            Assert.AreEqual(2, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void ListMarkerGetListMarkersShouldReturn1PersonListMarkerForTest1()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test1");
            var personListMarkers = ListWordMarker.GetListMarkers(ownerParts, "test1");

            Assert.AreEqual(1, personListMarkers.Count());
            Assert.AreEqual("el12[numId]2", personListMarkers.First().ElementId);
            Assert.AreEqual("el12[numId]2", personListMarkers.First().SelectionLastelementId);
        }

        [TestMethod]
        public void ListMarkerGetListMarkersShouldReturn1PersonListMarkerForTest2()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test2");
            var personListMarkers = ListWordMarker.GetListMarkers(ownerParts, "test2");

            Assert.AreEqual(1, personListMarkers.Count());
            Assert.AreEqual("el13[numId]2", personListMarkers.First().ElementId);
            Assert.AreEqual("el14[numId]2", personListMarkers.First().SelectionLastelementId);
        }

        [TestMethod]
        public void PictureMarkerGetSubdividedPartsShouldReturn2SelectorsForTest()
        {
            var markerWordSelectors = PictureDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[0]);

            Assert.AreEqual(2, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void PictureMarkerGetSubdividedPartsShouldReturn1SelectorForTest2()
        {
            var markerWordSelectors = PictureDocParagraphIdMarker.GetSubdividedParts(SplitDocumentObj.Person[1]);

            Assert.AreEqual(1, markerWordSelectors.Where(s => !string.IsNullOrEmpty(s.Email)).Count());
        }

        [TestMethod]
        public void PictureMarkerGetPictureMarkersShouldReturn1PictureMarkerForTest1()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test1");
            var personPictureMarkers = PictureWordMarker.GetPictureMarkers(ownerParts);

            Assert.AreEqual(1, personPictureMarkers.Count());
            Assert.AreEqual("el15", personPictureMarkers.First().ElementId);

        }

        [TestMethod]
        public void PictureMarkerGetPictureMarkersShouldReturn1PictureMarkerForTest2()
        {
            var ownerParts = PartsSelectionTreeElementMock.GetListMock().Where(p => p.Selected && p.OwnerName == "test2");
            var personPictureMarkers = PictureWordMarker.GetPictureMarkers(ownerParts);

            Assert.AreEqual(1, personPictureMarkers.Count());
            Assert.AreEqual("el16", personPictureMarkers.First().ElementId);
        }

        [TestCleanup]
        public void Finish()
        {
            WordNoParagraphIdDoc.Close();
            WordDemoDoc.Close();
        }
    }
}
