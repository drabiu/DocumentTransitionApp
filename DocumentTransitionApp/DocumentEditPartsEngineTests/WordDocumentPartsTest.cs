using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using DocumentEditPartsEngine.Interfaces;
using DocumentEditPartsEngine;
using System.Collections.Generic;

namespace DocumentEditPartsEngineTests
{
    [TestClass]
    public class WordDocumentPartsTest
    {
        MemoryStream WordNoParagraphIdDocInMemory;
        MemoryStream WordDemoDocInMemory;
        IDocumentParts WordDocumentParts;
        IList<PartsSelectionTreeElement> PartsSelectionElementsNoParagraphId;
        IList<PartsSelectionTreeElement> PartsSelectionElementsDemo;

        [TestInitialize]
        public void Init()
        {
            WordDocumentParts = DocumentPartsBuilder.Build(".docx");

            WordNoParagraphIdDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/lep-na-szczury-z-atr-karta-ch.docx"));
            WordDemoDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/demo.docx"));

            PartsSelectionElementsNoParagraphId = WordDocumentParts.Get(WordNoParagraphIdDocInMemory);
            PartsSelectionElementsDemo = WordDocumentParts.Get(WordDemoDocInMemory);
        }

        [TestMethod]
        public void WordGetMethodShouldReturn229Elements()
        {
            Assert.AreEqual(PartsSelectionElementsNoParagraphId.Count, 229);
        }

        [TestMethod]
        public void WordGetMethodShouldReturn90Elements()
        {
            Assert.AreEqual(PartsSelectionElementsDemo.Count, 90);
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectElementsNoParagraphId()
        {
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[10].ElementId, WordDocumentPartAttributes.GetParagraphNoIdFormatter(10));
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[100].ElementId, WordDocumentPartAttributes.GetParagraphNoIdFormatter(100));
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[200].ElementId, WordDocumentPartAttributes.GetParagraphNoIdFormatter(200));
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectElementsParagraphId()
        {
            Assert.AreEqual(PartsSelectionElementsDemo[1].ElementId, "2AD3D9AA");
            Assert.AreEqual(PartsSelectionElementsDemo[16].ElementId, "43F14223");
            Assert.AreEqual(PartsSelectionElementsDemo[24].ElementId, "6C45949E");
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectElementsName()
        {
            Assert.AreEqual(PartsSelectionElementsDemo[2].Name, "[Par]: There is support for images,");
            Assert.AreEqual(PartsSelectionElementsDemo[5].Name, "[Par]: Text Formatting");
            Assert.AreEqual(PartsSelectionElementsDemo[21].Name, "[Par]: Next, we have something a little");
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[2].Name, "[Par]: Nazwa handlowa Pułapka na szczury");
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[18].Name, "[Par]: Klasyfikacja produktu");
        }
     
        [TestCleanup]
        public void Finish()
        {
            WordNoParagraphIdDocInMemory.Close();
            WordDemoDocInMemory.Close();
        }
    }
}
