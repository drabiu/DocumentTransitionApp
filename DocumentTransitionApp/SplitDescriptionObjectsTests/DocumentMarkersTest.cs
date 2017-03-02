using Microsoft.VisualStudio.TestTools.UnitTesting;
using SplitDescriptionObjects;
using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace SplitDescriptionObjectsTests
{
    [TestClass]
    public class DocumentMarkersTest
    {
        IUniversalWordMarker UniversalDocNoParagraphIdMarker;
        IUniversalWordMarker UniversalDocParagraphIdMarker;
        WordprocessingDocument WordNoParagraphIdDoc;
        WordprocessingDocument WordDemoDoc;

        [TestInitialize]
        public void Init()
        {
            WordNoParagraphIdDoc = WordprocessingDocument.Open(@"../../../Files/lep-na-szczury-z-atr-karta-ch.docx", false);
            WordDemoDoc = WordprocessingDocument.Open(@"../../../Files/demo.docx", false);

            //test scenarios when paragraphs have an Id and a paragraph hasn`t got an Id
            UniversalDocNoParagraphIdMarker = new UniversalWordMarker(WordNoParagraphIdDoc.MainDocumentPart.Document.Body);
            UniversalDocParagraphIdMarker = new UniversalWordMarker(WordDemoDoc.MainDocumentPart.Document.Body);
        }

        [TestMethod]
        public void DocumentMarkerGetCrossedParagraphElementsShouldReturnOne()
        {
            IList<int> result = UniversalDocNoParagraphIdMarker.GetCrossedParagraphElements(WordDocumentPartAttributes.GetParagraphNoIdFormatter(1), WordDocumentPartAttributes.GetParagraphNoIdFormatter(1));

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(1, result[0]);
        }

        [TestMethod]
        public void DocumentMarkerGetCrossedParagraphElementsShouldReturnThree()
        {
            IList<int> result = UniversalDocParagraphIdMarker.GetCrossedParagraphElements("3CCBE53A", "4424AD34");

            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(17, result[0]);
            Assert.AreEqual(19, result[1]);
            Assert.AreEqual(20, result[2]);
        }

        [TestMethod]
        public void DocumentMarkerGetCrossedParagraphElementsShouldReturnNone()
        {
            IList<int> result = UniversalDocNoParagraphIdMarker.GetCrossedParagraphElements(WordDocumentPartAttributes.GetParagraphNoIdFormatter(4), WordDocumentPartAttributes.GetParagraphNoIdFormatter(1));

            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            WordNoParagraphIdDoc.Close();
            WordDemoDoc.Close();
        }
    }
}
