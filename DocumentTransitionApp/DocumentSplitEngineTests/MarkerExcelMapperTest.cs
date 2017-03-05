using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentSplitEngine.Excel;
using DocumentSplitEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class MarkerExcelMapperTest
    {
        IMarkerMapper<Sheet> MarkerExcelMapper;
        SpreadsheetDocument ExcelDemoDoc;

        [TestInitialize]
        public void Init()
        {
            byte[] sampleXmlBinary = File.ReadAllBytes(@"../../../Files/split_ExcelTutorialR1 — edytowalny.xlsx_20170304230639351.xml");

            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            Split splitXml = (Split)serializer.Deserialize(new MemoryStream(sampleXmlBinary));

            ExcelDemoDoc = SpreadsheetDocument.Open(@"../../../Files/ExcelTutorialR1 — edytowalny.xlsx", false);

            MarkerExcelMapper = new MarkerExcelMapper("ExcelTutorialR1 — edytowalny", splitXml, ExcelDemoDoc.WorkbookPart.Workbook);
        }

        [TestMethod]
        public void RunShouldReturn6Parts()
        {
            var documentPartList = MarkerExcelMapper.Run();

            Assert.AreEqual(7, documentPartList.Count);
        }

        [TestMethod]
        public void RunShouldReturn3ElementForPartByOwner()
        {
            var documentPartList = MarkerExcelMapper.Run();
            var ownerCompositeElements = documentPartList.SingleOrDefault(p => p.PartOwner == "test").CompositeElements;

            Assert.AreEqual(3, ownerCompositeElements.Count);
        }
        [TestMethod]
        public void RunShouldReturn1ElementForEachPartByOwner()
        {
            var documentPartList = MarkerExcelMapper.Run();
            var ownerCompositeElements = documentPartList.Where(p => p.PartOwner == "test2").Select(o => o.CompositeElements);

            Assert.AreEqual(2, ownerCompositeElements.Count());
            Assert.AreEqual(1, ownerCompositeElements.ElementAt(0).Count);
            Assert.AreEqual(1, ownerCompositeElements.ElementAt(1).Count);
        }

        [TestCleanup]
        public void Finish()
        {
            ExcelDemoDoc.Close();
        }
    }
}
