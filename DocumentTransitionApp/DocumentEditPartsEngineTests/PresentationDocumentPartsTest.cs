using DocumentEditPartsEngine;
using DocumentEditPartsEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentEditPartsEngineTests
{
    [TestClass]
    public class PresentationDocumentPartsTest
    {
        MemoryStream PreCGWDocInMemory;
        MemoryStream PreSampleDocInMemory;
        IPresentationParts PreDocumentParts;
        IList<PartsSelectionTreeElement> PartsSelectionElementsCGW;
        IList<PartsSelectionTreeElement> PartsSelectionElementsSample;

        [TestInitialize]
        public void Init()
        {
            PreDocumentParts = new PresentationDocumentParts();

            PreCGWDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/6.CGW15-prezentacja.pptx"));
            PreSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx"));

            PartsSelectionElementsCGW = PreDocumentParts.GetSlides(PreCGWDocInMemory);
            PartsSelectionElementsSample = PreDocumentParts.GetSlides(PreSampleDocInMemory);
        }

        [TestMethod]
        public void GetSlidesMethodShouldReturn18SlideElements()
        {
            Assert.AreEqual(18, PartsSelectionElementsCGW.Count);
        }

        [TestMethod]
        public void GetSlidesMethodShouldReturn13SlideElements()
        {
            Assert.AreEqual(13, PartsSelectionElementsSample.Count);
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsId()
        {
            Assert.AreEqual("rId4", PartsSelectionElementsCGW[1].ElementId);
            Assert.AreEqual("rId13", PartsSelectionElementsCGW[10].ElementId);
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsName()
        {
            Assert.AreEqual("[Sld]: Data processing in modern science", PartsSelectionElementsCGW[2].Name);
            Assert.AreEqual("[Sld]: Data processing in modern science", PartsSelectionElementsCGW[5].Name);
            Assert.AreEqual("[Sld]: Elementy slajdu", PartsSelectionElementsSample[3].Name);
            Assert.AreEqual("[Sld]: Grafika - Obrazy", PartsSelectionElementsSample[7].Name);
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsChildrenCount()
        {
            Assert.AreEqual(0, PartsSelectionElementsSample[3].Childs.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDocInMemory.Close();
            PreSampleDocInMemory.Close();
        }
    }
}
