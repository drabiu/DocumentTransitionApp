using DocumentEditPartsEngine;
using DocumentEditPartsEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLTools.Helpers;
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
        IDocumentParts PreDocumentParts;
        IList<PartsSelectionTreeElement> PartsSelectionElementsTutorial;
        IList<PartsSelectionTreeElement> PartsSelectionElementsTest;

        [TestInitialize]
        public void Init()
        {
            PreDocumentParts = new ExcelDocumentParts();

            ExcTutorialDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/ExcelTutorialR1 — edytowalny.xlsx"));
            ExcTestDocInMemory = new MemoryStream(File.ReadAllBytes(@"../../../Files/test.xlsx"));

            PartsSelectionElementsTutorial = PreDocumentParts.Get(ExcTutorialDocInMemory);
            PartsSelectionElementsTest = PreDocumentParts.Get(ExcTestDocInMemory);
        }

        [TestMethod]
        public void GetSheetsMethodShouldReturn6SheetElements()
        {
            var elementsTest = PartsSelectionElementsTest.Where(p => p.Type == ElementType.Sheet).ToList();

            Assert.AreEqual(6, elementsTest.Count);
        }

        [TestMethod]
        public void GetSheetsMethodShouldReturn82SheetElements()
        {
            var elementsTutorial = PartsSelectionElementsTutorial.Where(p => p.Type == ElementType.Sheet).ToList();

            Assert.AreEqual(82, elementsTutorial.Count);
        }

        [TestMethod]
        public void GetSheetsMethodShouldHaveCorrectSheetElementsId()
        {
            var elementsTest = PartsSelectionElementsTest.Where(p => p.Type == ElementType.Sheet).ToList();
            var elementsTutorial = PartsSelectionElementsTutorial.Where(p => p.Type == ElementType.Sheet).ToList();

            Assert.AreEqual("shId2", elementsTest[1].ElementId);
            Assert.AreEqual("shId11", elementsTutorial[10].ElementId);
        }

        [TestMethod]
        public void GetSheetsMethodShouldHaveCorrectSheetElementsName()
        {
            var elementsTest = PartsSelectionElementsTest.Where(p => p.Type == ElementType.Sheet).ToList();
            var elementsTutorial = PartsSelectionElementsTutorial.Where(p => p.Type == ElementType.Sheet).ToList();

            Assert.AreEqual("Hyperlink Test", elementsTest[1].Name);
            Assert.AreEqual("Notes", elementsTest[4].Name);
            Assert.AreEqual("Tabs, Ribbons", elementsTutorial[3].Name);
            Assert.AreEqual("Right-clicking", elementsTutorial[8].Name);
        }

        [TestMethod]
        public void GetSlidesMethodShouldHaveCorrectSlideElementsChildrenCount()
        {
            var elementsTutorial = PartsSelectionElementsTutorial.Where(p => p.Type == ElementType.Sheet).ToList();

            Assert.AreEqual(0, elementsTutorial[3].Childs.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            ExcTutorialDocInMemory.Close();
            ExcTestDocInMemory.Close();
        }
    }
}
