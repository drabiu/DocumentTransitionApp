using DocumentSplitEngine;
using DocumentSplitEngine.Interfaces;
using DocumentSplitEngineTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class ExcelSplitTest
    {
        int ErrorsCount;
        int WarningsCount;
        //MemoryStream ExcSampleDocInMemory;
        //MemoryStream ExcSampleDocXmlInMemory;
        ISplit ExcSampleSplit;
        ISplitXml SplitXml;

        byte[] SplitXmlBinary;
        XNamespace Xlmns = "https://sourceforge.net/p/documenttransitionapp/svn/HEAD/tree/DocumentTransitionApp/";

        [TestInitialize]
        public void Init()
        {
            var excelSplit = new ExcelSplit("test");
            ExcSampleSplit = excelSplit;
            SplitXml = excelSplit;

            var parts = PartsSelectionTreeElementMock.GetListMock();
            SplitXmlBinary = SplitXml.CreateSplitXml(parts);

            //ExcSampleDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/przykladowa-prezentacja.pptx"));
            //ExcSampleDocXmlInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/split_przykladowa-prezentacja.pptx_20170227215707619.xml"));
            ErrorsCount = 0;
            WarningsCount = 0;
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnValidXml()
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add(Xlmns.ToString(), @"../../../UnmarshallingSplitXml/splitXmlDefinitionTemplate.xsd");
            settings.ValidationType = ValidationType.Schema;

            XmlReader reader = XmlReader.Create(new MemoryStream(SplitXmlBinary), settings);
            XmlDocument document = new XmlDocument();
            document.Load(reader);
            document.Validate(ValidationEventHandler);

            XDocument xdoc = XDocument.Load(new MemoryStream(SplitXmlBinary));
            var elements = xdoc.Descendants(Xlmns + "Excel");

            Assert.AreEqual(0, ErrorsCount);
            Assert.AreEqual(0, WarningsCount);
            Assert.IsTrue(elements.Count() > 0);
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturn2PersonsWithCorrectEmails()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(SplitXmlBinary));
            var persons = xdoc.Descendants(Xlmns + "Person");
            var emails = persons.Select(el => el.Attribute("Email").Value);

            Assert.AreEqual(2, persons.Count());
            Assert.IsTrue(emails.Contains("test1"));
            Assert.IsTrue(emails.Contains("test2"));
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnPersonWith3UniversalMarkersAndProperSelection()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(SplitXmlBinary));
            var person = xdoc.Descendants(Xlmns + "Person").Where(el => el.Attribute("Email").Value == "test1");
            var markers = person.Elements(Xlmns + "UniversalMarker");


            Assert.AreEqual(3, markers.Count());
            Assert.AreEqual("el1", markers.ElementAt(0).Element(Xlmns + "ElementId").Value);
            Assert.AreEqual("el1", markers.ElementAt(0).Element(Xlmns + "SelectionLastelementId").Value);
            Assert.AreEqual("el2", markers.ElementAt(1).Element(Xlmns + "ElementId").Value);
            Assert.AreEqual("el2", markers.ElementAt(1).Element(Xlmns + "SelectionLastelementId").Value);
            Assert.AreEqual("el3", markers.ElementAt(2).Element(Xlmns + "ElementId").Value);
            Assert.AreEqual("el3", markers.ElementAt(2).Element(Xlmns + "SelectionLastelementId").Value);
        }

        [TestMethod]
        public void CreateSplitXMLShouldReturnPersonWith2UniversalMarkersAndProperSelection()
        {
            XDocument xdoc = XDocument.Load(new MemoryStream(SplitXmlBinary));
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
            var parts = SplitXml.SelectPartsFromSplitXml(new MemoryStream(SplitXmlBinary), PartsSelectionTreeElementMock.GetUnselectedPartsListMock());
            var selectedParts = parts.Where(p => p.Selected && p.OwnerName == "test1");

            Assert.AreEqual(3, selectedParts.Count());
        }

        [TestMethod]
        public void PartsFromSplitXMLShouldReturn2Selected()
        {
            var parts = SplitXml.SelectPartsFromSplitXml(new MemoryStream(SplitXmlBinary), PartsSelectionTreeElementMock.GetUnselectedPartsListMock());
            var selectedParts = parts.Where(p => p.Selected && p.OwnerName == "test2");

            Assert.AreEqual(2, selectedParts.Count());
        }

        [TestMethod]
        public void PartsFromSplitXMLShouldReturn2Unselected()
        {
            var parts = SplitXml.SelectPartsFromSplitXml(new MemoryStream(SplitXmlBinary), PartsSelectionTreeElementMock.GetUnselectedPartsListMock());
            var selectedParts = parts.Where(p => !p.Selected);

            Assert.AreEqual(2, selectedParts.Count());
        }

        [TestCleanup]
        public void Finish()
        {
            //ExcSampleDocInMemory.Close();
            //ExcSampleDocXmlInMemory.Close();
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
