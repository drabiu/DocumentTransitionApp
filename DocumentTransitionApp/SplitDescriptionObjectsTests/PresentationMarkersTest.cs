using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SplitDescriptionObjects;
using SplitDescriptionObjects.Interfaces;
using System.Collections.Generic;

namespace SplitDescriptionObjectsTests
{
    [TestClass]
    public class PresentationMarkersTest
    {
        ISlidePresentationMarker UniversalPreCGWMarker;
        ISlidePresentationMarker UniversalPreSampleMarker;
        PresentationDocument PreCGWDoc;
        PresentationDocument PreSampleDoc;

        [TestInitialize]
        public void Init()
        {
            PreCGWDoc = PresentationDocument.Open(@"../../../Files/6.CGW15-prezentacja.pptx", false);
            PreSampleDoc = PresentationDocument.Open(@"../../../Files/przykladowa-prezentacja.pptx", false);

            UniversalPreCGWMarker = new SlidePresentationMarker(PreCGWDoc.PresentationPart);
            UniversalPreSampleMarker = new SlidePresentationMarker(PreSampleDoc.PresentationPart);
        }

        [TestMethod]
        public void PresentationtMarkerGetCrossedSlideIdElementsShouldReturnOne()
        {
            IList<int> result = UniversalPreCGWMarker.GetCrossedSlideIdElements("rId5", "rId5");

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(2, result[0]);
        }

        [TestMethod]
        public void PresentationMarkerGetCrossedSlideIdElementsShouldReturnThree()
        {
            IList<int> result = UniversalPreSampleMarker.GetCrossedSlideIdElements("rId12", "rId14");

            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(10, result[0]);
            Assert.AreEqual(11, result[1]);
            Assert.AreEqual(12, result[2]);
        }

        [TestMethod]
        public void PresentationMarkerGetCrossedSlideIdElementsShouldReturnNone()
        {
            IList<int> result = UniversalPreCGWMarker.GetCrossedSlideIdElements("rId7", "rId5");

            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDoc.Close();
            PreSampleDoc.Close();
        }
    }
}
