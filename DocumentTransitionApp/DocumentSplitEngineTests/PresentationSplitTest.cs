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
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class PresentationSplitTest
    {
        int ErrorsCount;
        int WarningsCount;
        MemoryStream PreSampleDocInMemory;
        MemoryStream PreSampleDocXmlInMemory;
        ISplit DocSampleSplit;

        [TestInitialize]
        public void Init()
        {
            DocSampleSplit = new PresentationSplit("przykladowa-prezentacja");
            PreSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx"));
            PreSampleDocXmlInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/split_przykladowa-prezentacja.pptx_20170227215707619.xml"));
            ErrorsCount = 0;
            WarningsCount = 0;
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnValidXml()
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add("https://sourceforge.net/p/documenttransitionapp/svn/HEAD/tree/DocumentTransitionApp/", @"../../../UnmarshallingSplitXml/splitXmlDefinitionTemplate.xsd");
            settings.ValidationType = ValidationType.Schema;

            XmlReader reader = XmlReader.Create(@"../../../Files/split_przykladowa-prezentacja.pptx_20170227215707619.xml", settings);
            XmlDocument document = new XmlDocument();
            document.Load(reader);
            document.Validate(ValidationEventHandler);

            Assert.AreEqual(0, ErrorsCount);
            Assert.AreEqual(0, WarningsCount);
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturn2PersonsWithCorrectNames()
        {
            DocSampleSplit.OpenAndSearchDocument(PreSampleDocInMemory, PreSampleDocXmlInMemory);
            IList<PersonFiles> result = DocSampleSplit.SaveSplitDocument(PreSampleDocInMemory);
        }

        [TestCleanup]
        public void Finish()
        {
            PreSampleDocInMemory.Close();
            PreSampleDocXmlInMemory.Close();
        }

        private void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    ErrorsCount++;
                    break;
                case XmlSeverityType.Warning:
                    WarningsCount++;
                    break;
            }
        }
    }
}
