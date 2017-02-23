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
            Assert.Fail();
        }

        [TestMethod]
        public void ExcelMarkerGetCrossedSlideIdElementsShouldReturnThree()
        {
            Assert.Fail();
        }

        [TestMethod]
        public void ExcelMarkerGetCrossedSlideIdElementsShouldReturnNone()
        {
            Assert.Fail();
        }

        [TestCleanup]
        public void Finish()
        {
            ExcTutorialDoc.Close();
            ExcTestDoc.Close();
        }
    }
}
