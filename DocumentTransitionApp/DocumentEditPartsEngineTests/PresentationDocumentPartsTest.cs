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
            Assert.AreEqual(PartsSelectionElementsCGW.Count, 18);
        }

        [TestMethod]
        public void GetSlidesMethodShouldReturn13SlideElements()
        {
            Assert.AreEqual(PartsSelectionElementsSample.Count, 13);
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsId()
        {
            Assert.AreEqual(PartsSelectionElementsCGW[1].ElementId, "rId4");
            Assert.AreEqual(PartsSelectionElementsCGW[10].ElementId, "rId13");
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsName()
        {
            Assert.AreEqual(PartsSelectionElementsCGW[2].Name, "[Sld]: Data processing in modern science");
            Assert.AreEqual(PartsSelectionElementsCGW[5].Name, "[Sld]: Data processing in modern science");
            Assert.AreEqual(PartsSelectionElementsSample[3].Name, "[Sld]: Elementy slajdu");
            Assert.AreEqual(PartsSelectionElementsSample[7].Name, "[Sld]: Grafika - Obrazy");
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsChildrenCount()
        {
            Assert.AreEqual(PartsSelectionElementsSample[3].Childs.Count, 0);
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDocInMemory.Close();
            PreSampleDocInMemory.Close();
        }
    }
}
