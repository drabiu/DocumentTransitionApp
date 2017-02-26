using DocumentSplitEngine;
using DocumentSplitEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class PresentationSplitTest
    {
        MemoryStream PreCGWDocInMemory;
        MemoryStream PreSampleDocInMemory;
        MemoryStream PreCGWDocXmlInMemory;
        MemoryStream PreSampleDocXmlInMemory;
        ISplit DocCGWSplit;
        ISplit DocSampleSplit;

        [TestInitialize]
        public void Init()
        {
            DocCGWSplit = new PresentationSplit("6.CGW15-prezentacja");
            DocSampleSplit = new PresentationSplit("przykladowa-prezentacja");

            PreCGWDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/6.CGW15-prezentacja.pptx"));
            PreSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx"));

            PreCGWDocXmlInMemory = new MemoryStream();
            PreSampleDocXmlInMemory = new MemoryStream();          
        }

        [TestMethod]
        public void test()
        {
            DocCGWSplit.OpenAndSearchDocument(PreCGWDocInMemory, PreCGWDocXmlInMemory);
            IList<PersonFiles> result = DocCGWSplit.SaveSplitDocument(PreCGWDocInMemory);
        }

        [TestMethod]
        public void test1()
        {
            DocSampleSplit.OpenAndSearchDocument(PreSampleDocInMemory, PreSampleDocXmlInMemory);
            IList<PersonFiles> result = DocSampleSplit.SaveSplitDocument(PreSampleDocInMemory);
        }

        [TestCleanup]
        public void Finish()
        {
            PreCGWDocInMemory.Close();
            PreCGWDocXmlInMemory.Close();
            PreSampleDocInMemory.Close();
            PreSampleDocXmlInMemory.Close();
        }
    }
}
