using Microsoft.VisualStudio.TestTools.UnitTesting;
using SplitDescriptionObjects;
using DocumentEditPartsEngine;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace SplitDescriptionObjectsTests
{
    [TestClass]
    public class ExcelMarkersTest
    {
        IUniversalExcelMarker UniversalExcTutorialMarker;
        IUniversalExcelMarker UniversalExcTestMarker;
        SpreadsheetDocument ExcTutorialDoc;
        SpreadsheetDocument ExcTestDoc;

        [TestInitialize]
        public void Init()
        {
            ExcTutorialDoc = SpreadsheetDocument.Open(@"../../../Files/ExcelTutorialR1 — edytowalny.xlsx", false);
            ExcTestDoc = SpreadsheetDocument.Open(@"../../../Files/test.xlsx", false);

            UniversalExcTutorialMarker = new UniversalExcelMarker(ExcTutorialDoc.WorkbookPart.Workbook);
            UniversalExcTestMarker = new UniversalExcelMarker(ExcTestDoc.WorkbookPart.Workbook);
        }

        [TestMethod]
        public void ExcelMarkerGetCrossedSlideIdElementsShouldReturnOne()
        {
            IList<int> result = UniversalExcTestMarker.GetCrossedSheetElements("slId5", "slId5");

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(4, result[0]);
        }

        [TestMethod]
        public void ExcelMarkerGetCrossedSlideIdElementsShouldReturnThree()
        {
            IList<int> result = UniversalExcTutorialMarker.GetCrossedSheetElements("slId12", "slId14");

            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(11, result[0]);
            Assert.AreEqual(12, result[1]);
            Assert.AreEqual(13, result[2]);
        }

        [TestMethod]
        public void ExcelMarkerGetCrossedSlideIdElementsShouldReturnNone()
        {
            IList<int> result = UniversalExcTutorialMarker.GetCrossedSheetElements("slId7", "slId5");

            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestCleanup]
        public void Finish()
        {
            ExcTutorialDoc.Close();
            ExcTestDoc.Close();
        }
    }
}
