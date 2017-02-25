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

namespace OpenXMLToolsTests
{
    [TestClass]
    public class PresentationToolsTest
    {
        PresentationDocument PreCGWDoc;
        PresentationDocument PreSampleDoc;
        MemoryStream PreCGWDocInMemoryExpandable;
        MemoryStream PreSampleDocInMemoryExpandable;

        IPresentationTools PreTools;

        [TestInitialize]
        public void Init()
        {
            PreTools = new PresentationTools();

            MemoryStream PreCGWDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/6.CGW15-prezentacja.pptx"));
            MemoryStream PreSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx"));
            PreCGWDocInMemoryExpandable = new MemoryStream();
            PreSampleDocInMemoryExpandable = new MemoryStream();

            byte[] byteArray = StreamTools.ReadFully(PreCGWDocInMemory);
            PreCGWDocInMemoryExpandable.Write(byteArray, 0, byteArray.Length);

            byteArray = StreamTools.ReadFully(PreSampleDocInMemory);
            PreSampleDocInMemoryExpandable.Write(byteArray, 0, byteArray.Length);

            PreCGWDoc = PresentationDocument.Open(PreCGWDocInMemoryExpandable, true);
            PreSampleDoc = PresentationDocument.Open(PreSampleDocInMemoryExpandable, true);

            PreCGWDocInMemory.Close();
            PreSampleDocInMemory.Close();
        }

        [TestMethod]
        public void GetSlideTitleShouldReturnValidSlideTitle()
        {
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreCGWDoc, 1), 35);

            Assert.AreEqual(title, "[Sld]: Agenda");
        }

        [TestMethod]
        public void InsertSlideShouldInsertSlideToPosition3()
        {
            PresentationDocument document = PreTools.InsertNewSlide(PreCGWDoc, 3, "InsertSlideShouldInsertSlideToPosition3");
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(document, 3), 35);

            Assert.AreEqual(title, "[Sld]: InsertSlideShouldInsertSlideToPosition3");
        }

        [TestMethod]
        public void InsertSlideShouldInsertSlideToPosition4and3()
        {
            PreTools.InsertNewSlide(PreSampleDoc, 3, "InsertSlideShouldInsertSlideToPosition4");
            PreTools.InsertNewSlide(PreSampleDoc, 3, "InsertSlideShouldInsertSlideToPosition3");

            string title4 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreSampleDoc, 4), 35);
            string title3 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreSampleDoc, 3), 35);

            Assert.AreEqual(title4, "[Sld]: InsertSlideShouldInsertSlideToPosition4");
            Assert.AreEqual(title3, "[Sld]: InsertSlideShouldInsertSlideToPosition3");
        }

        [TestMethod]
        public void InsertSlideShouldInsertSlideToPosition3and8()
        {
            PreTools.InsertNewSlide(PreCGWDoc, 3, "InsertSlideShouldInsertSlideToPosition3");
            PreTools.InsertNewSlide(PreCGWDoc, 8, "InsertSlideShouldInsertSlideToPosition8");

            string title3 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreCGWDoc, 3), 35);
            string title8 = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(PreCGWDoc, 8), 35);

            Assert.AreEqual(title3, "[Sld]: InsertSlideShouldInsertSlideToPosition3");
            Assert.AreEqual(title8, "[Sld]: InsertSlideShouldInsertSlideToPosition8");
        }

        [TestMethod]
        public void InsertSlideShouldInsertSlideToLastPosition()
        {
            PresentationDocument document = PreTools.InsertNewSlide(PreSampleDoc, 13, "InsertSlideShouldInsertSlideToPosition13");
            string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(document, 13), 35);

            Assert.AreEqual(title, "[Sld]: InsertSlideShouldInsertSlideToPosition13");
        }

        [TestMethod]
        public void InsertSlideShouldNotInsertSlideToPositionOutOfRange()
        {
            try
            {
                PresentationDocument document = PreTools.InsertNewSlide(PreSampleDoc, 14, "InsertSlideShouldInsertSlideToPosition14");
                Assert.Fail();
            }
            catch(InvalidOperationException ex)
            {
                Assert.AreEqual(ex.Message, "The position is greather than number of slides");
            }           
        }

        [TestMethod]
        public void RemoveAllSlidesShouldResultInEmptyPresentation()
        {
            PresentationDocument document = PreTools.RemoveAllSlides(PreSampleDoc);
            var slideIdList = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>();

            Assert.AreEqual(slideIdList.Count(), 0);
            Assert.AreEqual(document.PresentationPart.SlideParts.Count(), 0);
        }

        [TestMethod]
        public void InsertSlideFromTemplateShouldAddValidSlide()
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument())
            {
                using (PresentationDocument output = streamDoc.GetPresentationDocument())
                {
                    //PreTools.InsertSlideFromTemplate(output, PreSampleDoc, "rId13");
                    output.Close();
                }

                streamDoc.GetModifiedDocument().SaveAs(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx");
            }

            //PresentationDocument document = PreTools.InsertSlideFromTemplate(PreCGWDoc, PreSampleDoc, "rId13");
            //var slideIdList = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>();
            //string title = PresentationTools.GetSlideTitle(PresentationTools.GetSlidePart(document, 18), 35);

            //Assert.AreEqual(slideIdList.Count(), 19);
            //Assert.AreEqual(document.PresentationPart.SlideParts.Count(), 19);
            //Assert.AreEqual(title, "[Sld]: Pokazy niestandardowe");

            //var byteArray = PreCGWDocInMemoryExpandable.ToArray();
            ////System.IO.File.WriteAllBytes(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx", mem.ToArray());
            //System.IO.File.WriteAllBytes(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx", byteArray);
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDoc.Close();
            PreSampleDoc.Close();
            PreCGWDocInMemoryExpandable.Close();
            PreSampleDocInMemoryExpandable.Close();
        }       
    }
}
