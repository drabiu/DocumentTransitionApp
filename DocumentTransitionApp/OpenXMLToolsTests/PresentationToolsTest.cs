﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OpenXMLTools.Interfaces;
using OpenXMLTools;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLToolsTests
{
    [TestClass]
    public class PresentationToolsTest
    {
        PresentationDocument PreCGWDoc;
        PresentationDocument PreSampleDoc;
        MemoryStream PreCGWDocInMemory;
        MemoryStream PreSampleDocInMemory;

        IPresentationTools PreTools;

        [TestInitialize]
        public void Init()
        {
            PreTools = new PresentationTools();

            PreCGWDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/6.CGW15-prezentacja.pptx"));
            PreSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx"));

            PreCGWDoc = PresentationDocument.Open(PreCGWDocInMemory, true);
            PreSampleDoc = PresentationDocument.Open(PreSampleDocInMemory, true);
        }

        [TestMethod]
        public void InsertSlideShouldInsertSlideToPosition3()
        {
            PresentationDocument document = PreTools.InsertNewSlide(PreCGWDoc, 3, "aaa");
            Assert.Fail();
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDocInMemory.Close();
            PreSampleDocInMemory.Close();
            PreCGWDoc.Close();
            PreSampleDoc.Close();
        }
    }
}
