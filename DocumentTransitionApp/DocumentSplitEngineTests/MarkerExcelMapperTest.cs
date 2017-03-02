using DocumentFormat.OpenXml.Packaging;
using DocumentSplitEngine.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class MarkerExcelMapperTest
    {
        IMarkerMapper<WorkbookPart> MarkerExcelMapper;
        SpreadsheetDocument ExcelDemoDoc;

        [TestInitialize]
        public void Init()
        {

        }

        [TestCleanup]
        public void Finish()
        { }
    }
}
