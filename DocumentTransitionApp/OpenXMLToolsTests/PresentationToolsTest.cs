using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OpenXMLTools.Interfaces;
using OpenXMLTools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.Linq;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Validation;

namespace OpenXMLToolsTests
{
    [TestClass]
    public class PresentationToolsTest
    {
        PresentationDocument PreCGWDoc;
        PresentationDocument PreSampleDoc;
        OpenXmlMemoryStreamDocument PreCGWDocInMemoryExpandable;
        OpenXmlMemoryStreamDocument PreSampleDocInMemoryExpandable;

        OpenXmlValidator DocValidator;
        IPresentationTools PreTools;

        [TestInitialize]
        public void Init()
        {
            PreTools = new PresentationTools();
            DocValidator = new OpenXmlValidator();

            byte[] PreCGWBytes = File.ReadAllBytes(@"../../../Files/6.CGW15-prezentacja.pptx");
            byte[] PreSampleBytes = File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx");

            MemoryStream PreCGWDocInMemory = new MemoryStream(PreCGWBytes, 0, PreCGWBytes.Length, true, true);
            MemoryStream PreSampleDocInMemory = new MemoryStream(PreSampleBytes, 0, PreSampleBytes.Length, true, true);

            OpenXmlPowerToolsDocument PreCGWDocPowerTools = new OpenXmlPowerToolsDocument("6.CGW15 - prezentacja.pptx", PreCGWDocInMemory);
            OpenXmlPowerToolsDocument PreSampleDocPowerTools = new OpenXmlPowerToolsDocument("6.CGW15 - przykladowa-prezentacja.pptx", PreSampleDocInMemory);

            PreCGWDocInMemoryExpandable = new OpenXmlMemoryStreamDocument(PreCGWDocPowerTools);
            PreSampleDocInMemoryExpandable = new OpenXmlMemoryStreamDocument(PreSampleDocPowerTools);

            PreCGWDoc = PreCGWDocInMemoryExpandable.GetPresentationDocument();
            PreSampleDoc = PreSampleDocInMemoryExpandable.GetPresentationDocument();

            PreCGWDocInMemory.Close();
            PreSampleDocInMemory.Close();
        }

        [TestMethod]
        public void GetSlideTitleShouldReturnAccurateSlideTitle()
        {
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreCGWDoc, 1), 35);

            Assert.AreEqual("[Sld]: Agenda", title);
        }

        [TestMethod]
        public void InsertNewSlideShouldInsertSlideToPosition3()
        {
            PresentationDocument document = PreTools.InsertNewSlide(PreCGWDoc, 3, "InsertSlideShouldInsertSlideToPosition3");
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(document, 3), 35);

            Assert.AreEqual("[Sld]: InsertSlideShouldInsertSlideToPosition3", title);
        }

        [TestMethod]
        public void InsertNewSlideShouldInsertSlideToPosition4and3()
        {
            PreTools.InsertNewSlide(PreSampleDoc, 3, "InsertSlideShouldInsertSlideToPosition4");
            PreTools.InsertNewSlide(PreSampleDoc, 3, "InsertSlideShouldInsertSlideToPosition3");

            string title4 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreSampleDoc, 4), 35);
            string title3 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreSampleDoc, 3), 35);

            Assert.AreEqual("[Sld]: InsertSlideShouldInsertSlideToPosition4", title4);
            Assert.AreEqual("[Sld]: InsertSlideShouldInsertSlideToPosition3", title3);
        }

        [TestMethod]
        public void InsertNewSlideShouldInsertSlideToPosition3and8()
        {
            PreTools.InsertNewSlide(PreCGWDoc, 3, "InsertSlideShouldInsertSlideToPosition3");
            PreTools.InsertNewSlide(PreCGWDoc, 8, "InsertSlideShouldInsertSlideToPosition8");

            string title3 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreCGWDoc, 3), 35);
            string title8 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreCGWDoc, 8), 35);

            Assert.AreEqual("[Sld]: InsertSlideShouldInsertSlideToPosition3", title3);
            Assert.AreEqual("[Sld]: InsertSlideShouldInsertSlideToPosition8", title8);
        }

        [TestMethod]
        public void InsertNewSlideShouldInsertSlideToLastPosition()
        {
            PresentationDocument document = PreTools.InsertNewSlide(PreSampleDoc, 13, "InsertSlideShouldInsertSlideToPosition13");
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(document, 13), 35);

            Assert.AreEqual("[Sld]: InsertSlideShouldInsertSlideToPosition13", title);
        }

        [TestMethod]
        public void InsertNewSlideShouldNotInsertSlideToPositionOutOfRange()
        {
            try
            {
                PresentationDocument document = PreTools.InsertNewSlide(PreSampleDoc, 14, "InsertSlideShouldInsertSlideToPosition14");
                Assert.Fail();
            }
            catch(InvalidOperationException ex)
            {
                Assert.AreEqual("The position is greather than number of slides", ex.Message);
            }           
        }

        [TestMethod]
        public void InsertNewSlideShouldResultInValidDocument()
        {
            PresentationDocument document = PreTools.InsertNewSlide(PreCGWDoc, 3, "InsertSlideShouldInsertSlideToPosition3");
            var validationErrors = DocValidator.Validate(document);

            Assert.AreEqual(0, validationErrors.Count());
        }

        [TestMethod]
        public void RemoveAllSlidesShouldResultInEmptyPresentation()
        {
            PresentationDocument document = PreTools.RemoveAllSlides(PreSampleDoc);
            var slideIdList = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>();

            Assert.AreEqual(0, slideIdList.Count());
            Assert.AreEqual(0, document.PresentationPart.SlideParts.Count());
        }

        [TestMethod]
        public void RemoveAllSlidesShouldResultInValidDocument()
        {
            PresentationDocument document = PreTools.RemoveAllSlides(PreSampleDoc);
            var validationErrors = DocValidator.Validate(document);

            Assert.AreEqual(0, validationErrors.Count());
        }

        [TestMethod]
        public void InsertSlideFromTemplateShouldAddValidSlide()
        {
            PresentationDocument document = PreTools.InsertSlideFromTemplate(PreCGWDoc, PreSampleDoc, "rId13");
            var slideIdList = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>();
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(document, 18), 35);
            var validationErrors = DocValidator.Validate(document);

            Assert.AreEqual(19, slideIdList.Count());
            Assert.AreEqual(19, document.PresentationPart.SlideParts.Count());
            Assert.AreEqual("[Sld]: Pokazy niestandardowe", title);
            Assert.AreEqual(0, validationErrors.Count());

            PreCGWDocInMemoryExpandable.GetModifiedDocument().SaveAs(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx");
        }

        [TestCleanup]
        public void Finish()
        {
            //PreCGWDoc.Close();
            PreSampleDoc.Close();
            PreCGWDocInMemoryExpandable.Close();
            PreSampleDocInMemoryExpandable.Close();
        }       
    }
}
