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

            PartsSelectionElementsTutorial = PreDocumentParts.Get(ExcTutorialDocInMemory);
            PartsSelectionElementsTest = PreDocumentParts.Get(ExcTestDocInMemory);
        }

        [TestCleanup]
        public void Finish()
        {
            ExcTutorialDocInMemory.Close();
            ExcTestDocInMemory.Close();
        }
    }
}
