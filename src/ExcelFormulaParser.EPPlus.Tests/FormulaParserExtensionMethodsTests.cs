using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.EPPlus;
using System.IO;

namespace ExcelFormulaParser.EPPlus.Tests
{
    [TestClass]
    public class FormulaParserExtensionMethodsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _workSheet;


        [TestInitialize]
        public void Initialize()
        {
            var ms = new MemoryStream();
            _package = new ExcelPackage(ms);
            _workSheet = _package.Workbook.Worksheets.Add("Test");
            for (var x = 1; x < 11; x++)
            {
                _workSheet.Cells["A" + x].Formula = "(1 + 2) * 2";
            }
            _workSheet.Cells["A12"].Formula = "SUM(A1:A10)";
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void WorksheetShouldCalculateAll()
        {
            _workSheet.Calculate(_package);
            Assert.AreEqual(60d, _workSheet.Cells["A12"].Value);
        }
    }
}
