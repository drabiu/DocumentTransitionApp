using DocumentEditPartsEngine;
using DocumentEditPartsEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngineTests
{
    [TestClass]
    public class ExcelDocumentPartsTest
    {
        MemoryStream ExcTutorialDocInMemory;
        MemoryStream ExcTestDocInMemory;
        IExcelParts PreDocumentParts;
        IList<PartsSelectionTreeElement> PartsSelectionElementsTutorial;
        IList<PartsSelectionTreeElement> PartsSelectionElementsTest;

        [TestInitialize]
        public void Init()
        {
            PreDocumentParts = new ExcelDocumentParts();

            ExcTutorialDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/ExcelTutorialR1 — edytowalny.xlsx"));
            ExcTestDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/test.xlsx"));

            PartsSelectionElementsTutorial = PreDocumentParts.GetSheets(ExcTutorialDocInMemory);
            PartsSelectionElementsTest = PreDocumentParts.GetSheets(ExcTestDocInMemory);
        }

        [TestMethod]
        public void GetSheetsMethodShouldReturn6SheetElements()
        {
            Assert.AreEqual(6, PartsSelectionElementsTest.Count);
        }

        [TestMethod]
        public void GetSheetsMethodShouldReturn82SheetElements()
        {
            Assert.AreEqual(82, PartsSelectionElementsTutorial.Count);
        }

        [TestMethod]
        public void GetSheetsMethodShouldHaveCorrectSheetElementsId()
        {
            Assert.AreEqual("slId2", PartsSelectionElementsTest[1].ElementId);
            Assert.AreEqual("slId11", PartsSelectionElementsTutorial[10].ElementId);
        }

        [TestMethod]
        public void GetSheetsMethodShouldHaveCorrectSheetElementsName()
        {
            Assert.AreEqual("[Sht]: Hyperlink Test", PartsSelectionElementsTest[1].Name);
            Assert.AreEqual("[Sht]: Notes", PartsSelectionElementsTest[4].Name);
            Assert.AreEqual("[Sht]: Tabs, Ribbons", PartsSelectionElementsTutorial[3].Name);
            Assert.AreEqual("[Sht]: Right-clicking", PartsSelectionElementsTutorial[8].Name);
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsChildrenCount()
        {
            Assert.AreEqual(0, PartsSelectionElementsTutorial[3].Childs.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            ExcTutorialDocInMemory.Close();
            ExcTestDocInMemory.Close();
        }
    }
}
