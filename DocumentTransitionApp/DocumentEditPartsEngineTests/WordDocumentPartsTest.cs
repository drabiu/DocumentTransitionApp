using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using DocumentEditPartsEngine.Interfaces;
using DocumentEditPartsEngine;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentEditPartsEngineTests
{
    [TestClass]
    public class WordDocumentPartsTest
    {
        MemoryStream WordNoParagraphIdDocInMemory;
        MemoryStream WordDemoDocInMemory;
        IDocumentParts WordDocumentParts;
        IList<PartsSelectionTreeElement> PartsSelectionElementsNoParagraphId;
        IList<PartsSelectionTreeElement> PartsSelectionElementsParagraphDemo;

        [TestInitialize]
        public void Init()
        {
            WordDocumentParts = new WordDocumentParts();

            WordNoParagraphIdDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/lep-na-szczury-z-atr-karta-ch.docx"));
            WordDemoDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/demo.docx"));

            PartsSelectionElementsNoParagraphId = WordDocumentParts.Get(WordNoParagraphIdDocInMemory, el => el is Paragraph);
            PartsSelectionElementsParagraphDemo = WordDocumentParts.Get(WordDemoDocInMemory, el => el is Paragraph);
        }

        [TestMethod]
        public void WordGetMethodShouldReturn229ParagraphElements()
        {
            Assert.AreEqual(PartsSelectionElementsNoParagraphId.Count, 229);
        }

        [TestMethod]
        public void WordGetMethodShouldReturn90ParagraphElements()
        {
            Assert.AreEqual(PartsSelectionElementsParagraphDemo.Count, 90);
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
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[1].ElementId, "2AD3D9AA");
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[16].ElementId, "43F14223");
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[24].ElementId, "6C45949E");
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectParagraphElementsName()
        {
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[2].Name, "[Par]: There is support for images,");
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[5].Name, "[Par]: Text Formatting");
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[21].Name, "[Par]: Next, we have something a little");
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[2].Name, "[Par]: Nazwa handlowa Pułapka na szczury");
            Assert.AreEqual(PartsSelectionElementsNoParagraphId[18].Name, "[Par]: Klasyfikacja produktu");
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectParagraphElementsChildrenCount()
        {
            Assert.AreEqual(PartsSelectionElementsParagraphDemo[3].Childs.Count, 0);
        }

        [TestCleanup]
        public void Finish()
        {
            WordNoParagraphIdDocInMemory.Close();
            WordDemoDocInMemory.Close();
        }
    }
}
