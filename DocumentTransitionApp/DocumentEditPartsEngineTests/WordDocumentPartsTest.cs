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
            Assert.AreEqual(229, PartsSelectionElementsNoParagraphId.Count);
        }

        [TestMethod]
        public void WordGetMethodShouldReturn90ParagraphElements()
        {
            Assert.AreEqual(90, PartsSelectionElementsParagraphDemo.Count);
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectElementsNoParagraphId()
        {
            Assert.AreEqual(WordDocumentPartAttributes.GetParagraphNoIdFormatter(10), PartsSelectionElementsNoParagraphId[10].ElementId);
            Assert.AreEqual(WordDocumentPartAttributes.GetParagraphNoIdFormatter(100), PartsSelectionElementsNoParagraphId[100].ElementId);
            Assert.AreEqual(WordDocumentPartAttributes.GetParagraphNoIdFormatter(200), PartsSelectionElementsNoParagraphId[200].ElementId);
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectElementsParagraphId()
        {
            Assert.AreEqual("2AD3D9AA", PartsSelectionElementsParagraphDemo[1].ElementId);
            Assert.AreEqual("43F14223", PartsSelectionElementsParagraphDemo[16].ElementId);
            Assert.AreEqual("6C45949E", PartsSelectionElementsParagraphDemo[24].ElementId);
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectParagraphElementsName()
        {
            Assert.AreEqual("[Par]: There is support for images,", PartsSelectionElementsParagraphDemo[2].Name);
            Assert.AreEqual("[Par]: Text Formatting", PartsSelectionElementsParagraphDemo[5].Name);
            Assert.AreEqual("[Par]: Next, we have something a little", PartsSelectionElementsParagraphDemo[21].Name);
            Assert.AreEqual("[Par]: Nazwa handlowa Pułapka na szczury", PartsSelectionElementsNoParagraphId[2].Name);
            Assert.AreEqual("[Par]: Klasyfikacja produktu", PartsSelectionElementsNoParagraphId[18].Name);
        }

        [TestMethod]
        public void WordGetMethodShouldHaveCorrectParagraphElementsChildrenCount()
        {
            Assert.AreEqual(0, PartsSelectionElementsParagraphDemo[3].Childs.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            WordNoParagraphIdDocInMemory.Close();
            WordDemoDocInMemory.Close();
        }
    }
}
