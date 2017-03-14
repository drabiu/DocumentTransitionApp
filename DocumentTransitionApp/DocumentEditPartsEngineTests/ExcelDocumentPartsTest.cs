using DocumentEditPartsEngine;
using DocumentEditPartsEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
            var elementsTest = PartsSelectionElementsTest.Where(p => p.Type == DocumentEditPartsEngine.Helpers.ElementType.Sheet).ToList();
            var elementsTutorial = PartsSelectionElementsTutorial.Where(p => p.Type == DocumentEditPartsEngine.Helpers.ElementType.Sheet).ToList();

            Assert.AreEqual("slId2", elementsTest[1].ElementId);
            Assert.AreEqual("slId11", elementsTutorial[10].ElementId);
        }

        [TestMethod]
        public void GetSheetsMethodShouldHaveCorrectSheetElementsName()
        {
            Assert.AreEqual("Hyperlink Test", PartsSelectionElementsTest[1].Name);
            Assert.AreEqual("Notes", PartsSelectionElementsTest[4].Name);
            Assert.AreEqual("Tabs, Ribbons", PartsSelectionElementsTutorial[3].Name);
            Assert.AreEqual("Right-clicking", PartsSelectionElementsTutorial[8].Name);
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
