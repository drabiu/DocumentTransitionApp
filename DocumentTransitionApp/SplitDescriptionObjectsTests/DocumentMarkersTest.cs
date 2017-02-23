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
        IUniversalDocumentMarker UniversalDocNoParagraphIdMarker;
        IUniversalDocumentMarker UniversalDocParagraphIdMarker;
        WordprocessingDocument WordNoParagraphIdDoc;
        WordprocessingDocument WordDemoDoc;

        [TestInitialize]
        public void Init()
        {
            WordNoParagraphIdDoc = WordprocessingDocument.Open(@"../../../Files/lep-na-szczury-z-atr-karta-ch.docx", false);
            WordDemoDoc = WordprocessingDocument.Open(@"../../../Files/demo.docx", false);

            //test scenarios when paragraphs have an Id and a paragraph hasn`t got an Id
            UniversalDocNoParagraphIdMarker = new UniversalDocumentMarker(WordNoParagraphIdDoc.MainDocumentPart.Document.Body);
            UniversalDocParagraphIdMarker = new UniversalDocumentMarker(WordDemoDoc.MainDocumentPart.Document.Body);
        }

        [TestMethod]
        public void DocumentMarkerGetCrossedParagraphElementsShouldReturnOne()
        {
            IList<int> result = UniversalDocNoParagraphIdMarker.GetCrossedParagraphElements(WordDocumentPartAttributes.GetParagraphNoIdFormatter(1), WordDocumentPartAttributes.GetParagraphNoIdFormatter(1));

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 1);
            Assert.AreEqual(result[0], 1);
        }

        [TestMethod]
        public void DocumentMarkerGetCrossedParagraphElementsShouldReturnThree()
        {
            IList<int> result = UniversalDocParagraphIdMarker.GetCrossedParagraphElements("3CCBE53A", "4424AD34");

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 3);
            Assert.AreEqual(result[0], 17);
            Assert.AreEqual(result[1], 19);
            Assert.AreEqual(result[2], 20);
        }

        [TestMethod]
        public void DocumentMarkerGetCrossedParagraphElementsShouldReturnNone()
        {
            IList<int> result = UniversalDocNoParagraphIdMarker.GetCrossedParagraphElements(WordDocumentPartAttributes.GetParagraphNoIdFormatter(4), WordDocumentPartAttributes.GetParagraphNoIdFormatter(1));

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 0);
        }

        [TestCleanup]
        public void Finish()
        {
            WordNoParagraphIdDoc.Close();
            WordDemoDoc.Close();
        }
    }
}
