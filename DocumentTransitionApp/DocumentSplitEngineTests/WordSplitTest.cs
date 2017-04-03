using DocumentFormat.OpenXml.Validation;
using DocumentSplitEngine;
using DocumentSplitEngine.Interfaces;
using DocumentSplitEngineTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class WordSplitTest
    {
        int ErrorsCount;
        int WarningsCount;
        MemoryStream WordSampleDocInMemory;

        ISplit WordSampleSplit;
        ISplitXml SplitXml;
        OpenXmlValidator DocValidator;

        //testing merge since it`s abstract
        IMergeXml WordSampleMerge;

        byte[] CreateSplitXmlBinary;
        byte[] MergeXmlBinary;
        XNamespace Xlmns = "https://sourceforge.net/p/documenttransitionapp/svn/HEAD/tree/DocumentTransitionApp/";

        [TestInitialize]
        public void Init()
        {
            var wordSplit = new WordSplit("demo");
            WordSampleMerge = wordSplit;
            WordSampleSplit = wordSplit;
            SplitXml = wordSplit;

            DocValidator = new OpenXmlValidator();

            var parts = PartsSelectionTreeElementMock.GetListMock();
            CreateSplitXmlBinary = SplitXml.CreateSplitXml(parts);

            WordSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/demo.docx"));

            byte[] sampleXmlBinary = File.ReadAllBytes(@"../../../Files/split_demo.docx_20170227215840894.xml");
            wordSplit.OpenAndSearchDocument(WordSampleDocInMemory, new MemoryStream(sampleXmlBinary));

            MergeXmlBinary = WordSampleMerge.CreateMergeXml();

            ErrorsCount = 0;
            WarningsCount = 0;
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnValidXml()
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add(Xlmns.ToString(), @"../../../UnmarshallingSplitXml/splitXmlDefinitionTemplate.xsd");
            settings.ValidationType = ValidationType.Schema;

            XmlReader reader = XmlReader.Create(new MemoryStream(CreateSplitXmlBinary), settings);
            XmlDocument document = new XmlDocument();
            document.Load(reader);
            document.Validate(ValidationEventHandler);

            XDocument xdoc = XDocument.Load(new MemoryStream(CreateSplitXmlBinary));
            var elements = xdoc.Descendants(Xlmns + "Document");

            Assert.AreEqual(0, ErrorsCount);
            Assert.AreEqual(0, WarningsCount);
            Assert.IsTrue(elements.Count() > 0);
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturn2PersonsWithCorrectEmails()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(CreateSplitXmlBinary));
            var persons = xdoc.Descendants(Xlmns + "Person");
            var emails = persons.Select(el => el.Attribute("Email").Value);

            Assert.AreEqual(2, persons.Count());
            Assert.IsTrue(emails.Contains("test1"));
            Assert.IsTrue(emails.Contains("test2"));
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnPersonWith3UniversalMarkersAndProperSelection()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(CreateSplitXmlBinary));
            var person = xdoc.Descendants(Xlmns + "Person").Where(el => el.Attribute("Email").Value == "test1");
            var markers = person.Elements(Xlmns + "UniversalMarker");

            Assert.AreEqual(1, markers.Count());
            Assert.AreEqual("el1", markers.ElementAt(0).Element(Xlmns + "ElementId").Value);
            Assert.AreEqual("el3", markers.ElementAt(0).Element(Xlmns + "SelectionLastelementId").Value);
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnPersonWith2UniversalMarkersAndProperSelection()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(CreateSplitXmlBinary));
            var person = xdoc.Descendants(Xlmns + "Person").Where(el => el.Attribute("Email").Value == "test2");
            var markers = person.Elements(Xlmns + "UniversalMarker");

            Assert.AreEqual(2, markers.Count());
            Assert.AreEqual("el5", markers.ElementAt(0).Element(Xlmns + "ElementId").Value);
            Assert.AreEqual("el5", markers.ElementAt(0).Element(Xlmns + "SelectionLastelementId").Value);
            Assert.AreEqual("el7", markers.ElementAt(1).Element(Xlmns + "ElementId").Value);
            Assert.AreEqual("el7", markers.ElementAt(1).Element(Xlmns + "SelectionLastelementId").Value);
        }

        [TestMethod]
        public void PartsFromSplitXMLShouldReturn3Selected()
        {
            var parts = SplitXml.SelectPartsFromSplitXml(new MemoryStream(CreateSplitXmlBinary), PartsSelectionTreeElementMock.GetUnselectedPartsListMock());
            var selectedParts = parts.Where(p => p.Selected && p.OwnerName == "test1");

            Assert.AreEqual(3, selectedParts.Count());
        }

        [TestMethod]
        public void PartsFromSplitXMLShouldReturn2Selected()
        {
            var parts = SplitXml.SelectPartsFromSplitXml(new MemoryStream(CreateSplitXmlBinary), PartsSelectionTreeElementMock.GetUnselectedPartsListMock());
            var selectedParts = parts.Where(p => p.Selected && p.OwnerName == "test2");

            Assert.AreEqual(2, selectedParts.Count());
        }

        [TestMethod]
        public void PartsFromSplitXMLShouldReturn2Unselected()
        {
            var parts = SplitXml.SelectPartsFromSplitXml(new MemoryStream(CreateSplitXmlBinary), PartsSelectionTreeElementMock.GetUnselectedPartsListMock());
            var selectedParts = parts.Where(p => !p.Selected);

            Assert.AreEqual(2, selectedParts.Count());
        }

        [TestMethod]
        public void CreateMergeXMLShouldReturnValidXml()
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add(Xlmns.ToString(), @"../../../UnmarshallingSplitXml/mergeXmlDefinitionTemplate.xsd");
            settings.ValidationType = ValidationType.Schema;

            XmlReader reader = XmlReader.Create(new MemoryStream(MergeXmlBinary), settings);
            XmlDocument document = new XmlDocument();
            document.Load(reader);
            document.Validate(ValidationEventHandler);

            XDocument xdoc = XDocument.Load(new MemoryStream(MergeXmlBinary));
            var elements = xdoc.Descendants(Xlmns + "Document");

            Assert.AreEqual(0, ErrorsCount);
            Assert.AreEqual(0, WarningsCount);
            Assert.IsTrue(elements.Count() > 0);
        }

        [TestMethod]
        public void CreateMergeXMLShouldReturn7Parts()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(MergeXmlBinary));
            var parts = xdoc.Descendants(Xlmns + "Part");

            Assert.AreEqual(7, parts.Count());
        }

        [TestMethod]
        public void CreateMergeXMLShouldReturn4UndefinedParts()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(MergeXmlBinary));
            var parts = xdoc.Descendants(Xlmns + "Part").Where(el => el.Element(Xlmns + "Name").Value == "undefined");

            Assert.AreEqual(4, parts.Count());
        }

        [TestMethod]
        public void CreateMergeXMLShouldReturn2Test2Parts()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(MergeXmlBinary));
            var parts = xdoc.Descendants(Xlmns + "Part").Where(el => el.Element(Xlmns + "Name").Value == "test2");

            Assert.AreEqual(2, parts.Count());
        }

        [TestMethod]
        public void CreateMergeXMLShouldReturn1TestParts()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(MergeXmlBinary));
            var parts = xdoc.Descendants(Xlmns + "Part").Where(el => el.Element(Xlmns + "Name").Value == "test");

            Assert.AreEqual(1, parts.Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturn9PersonFiles()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);

            Assert.AreEqual(9, personFilesList.Count);
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturn4Undefined()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);

            Assert.AreEqual(4, personFilesList.Where(p => p.Person == "undefined").Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturn1Test()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);

            Assert.AreEqual(1, personFilesList.Where(p => p.Person == "test").Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturn2Test2()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);

            Assert.AreEqual(2, personFilesList.Where(p => p.Person == "test2").Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturnTemplate()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);

            Assert.AreEqual(1, personFilesList.Where(p => p.Person == "/" && p.Name == "template.docx").Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturnMergeXml()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);

            Assert.AreEqual(1, personFilesList.Where(p => p.Person == "/" && p.Name == "mergeXmlDefinition.xml").Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturnValidUndefinedDocuments()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);
            var docs = personFilesList.Where(p => p.Person == "undefined").Select(u => u.Data);

            List<ValidationErrorInfo> validationErrors = new List<ValidationErrorInfo>();
            foreach (byte[] doc in docs)
            {
                MemoryStream partDocInMemory = new MemoryStream(doc, 0, doc.Length, true, true);
                var partDocPowerTools = new OpenXmlPowerToolsDocument("undefined.docx", partDocInMemory);

                OpenXmlMemoryStreamDocument partDocInMemoryExpandable = new OpenXmlMemoryStreamDocument(partDocPowerTools);

                validationErrors.AddRange(DocValidator.Validate(partDocInMemoryExpandable.GetWordprocessingDocument()));
            }

            Assert.IsTrue(docs.Count() > 0);
            Assert.AreEqual(0, validationErrors.Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturnValidTestDocuments()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);
            var docs = personFilesList.Where(p => p.Person == "test").Select(u => u.Data);

            List<ValidationErrorInfo> validationErrors = new List<ValidationErrorInfo>();
            foreach (byte[] doc in docs)
            {
                MemoryStream partDocInMemory = new MemoryStream(doc, 0, doc.Length, true, true);
                var partDocPowerTools = new OpenXmlPowerToolsDocument("test.docx", partDocInMemory);

                OpenXmlMemoryStreamDocument partDocInMemoryExpandable = new OpenXmlMemoryStreamDocument(partDocPowerTools);

                validationErrors.AddRange(DocValidator.Validate(partDocInMemoryExpandable.GetWordprocessingDocument()));
            }

            Assert.IsTrue(docs.Count() > 0);
            Assert.AreEqual(0, validationErrors.Count());
        }

        [TestMethod]
        public void SaveSplitDocumentShouldReturnValidTest2dDocuments()
        {
            var personFilesList = WordSampleSplit.SaveSplitDocument(WordSampleDocInMemory);
            var docs = personFilesList.Where(p => p.Person == "test2").Select(u => u.Data);

            List<ValidationErrorInfo> validationErrors = new List<ValidationErrorInfo>();
            foreach (byte[] doc in docs)
            {
                MemoryStream partDocInMemory = new MemoryStream(doc, 0, doc.Length, true, true);
                var partDocPowerTools = new OpenXmlPowerToolsDocument("test2.docx", partDocInMemory);

                OpenXmlMemoryStreamDocument partDocInMemoryExpandable = new OpenXmlMemoryStreamDocument(partDocPowerTools);

                validationErrors.AddRange(DocValidator.Validate(partDocInMemoryExpandable.GetWordprocessingDocument()));
            }

            Assert.IsTrue(docs.Count() > 0);
            Assert.AreEqual(0, validationErrors.Count());
        }

        [TestCleanup]
        public void Finish()
        {
            WordSampleDocInMemory.Close();
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
