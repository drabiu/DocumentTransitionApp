using Microsoft.VisualStudio.TestTools.UnitTesting;
using SplitDescriptionObjects;
using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace SplitDescriptionObjectsTests
{
    [TestClass]
    public class PresentationMarkersTest
    {
        IUniversalPresentationMarker UniversalPreCGWMarker;
        IUniversalPresentationMarker UniversalPreSampleMarker;
        PresentationDocument PreCGWDoc;
        PresentationDocument PreSampleDoc;

        [TestInitialize]
        public void Init()
        {
            PreCGWDoc = PresentationDocument.Open(@"../../../Files/6.CGW15-prezentacja.pptx", false);
            PreSampleDoc = PresentationDocument.Open(@"../../../Files/przykladowa-prezentacja.pptx", false);

            UniversalPreCGWMarker = new UniversalPresentationMarker(PreCGWDoc.PresentationPart);
            UniversalPreSampleMarker = new UniversalPresentationMarker(PreSampleDoc.PresentationPart);
        }

        [TestMethod]
        public void PresentationtMarkerGetCrossedSlideIdElementsShouldReturnOne()
        {
            IList<int> result = UniversalPreCGWMarker.GetCrossedSlideIdElements("rId5", "rId5");

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 1);
            Assert.AreEqual(result[0], 2);
        }

        [TestMethod]
        public void PresentationMarkerGetCrossedSlideIdElementsShouldReturnThree()
        {
            IList<int> result = UniversalPreSampleMarker.GetCrossedSlideIdElements("rId12", "rId14");

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 3);
            Assert.AreEqual(result[0], 10);
            Assert.AreEqual(result[1], 11);
            Assert.AreEqual(result[2], 12);
        }

        [TestMethod]
        public void PresentationMarkerGetCrossedSlideIdElementsShouldReturnNone()
        {
            IList<int> result = UniversalPreCGWMarker.GetCrossedSlideIdElements("rId7", "rId5");

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 0);
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDoc.Close();
            PreSampleDoc.Close();
        }
    }
}
