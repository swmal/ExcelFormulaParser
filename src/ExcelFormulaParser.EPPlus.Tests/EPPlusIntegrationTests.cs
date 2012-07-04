using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.EPPlus.Tests
{
    [TestClass]
    public class EPPlusIntegrationTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _workSheet;
        private FormulaParser _parser;


        [TestInitialize]
        public void Initialize()
        {
            var ms = new MemoryStream();
            _package = new ExcelPackage(ms);
            _workSheet = _package.Workbook.Worksheets.Add("Test");
            for (var x = 1; x < 11; x++)
            {
                _workSheet.Cells["A" + x].Value = x;
            }
            var dataProvider = new EPPlusExcelDataProvider(_package);
            _parser = new FormulaParser(dataProvider);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldRetrieveRangeValuesFromEPPlus()
        {
            var result = _parser.Parse("=SUM(A1:A5) - SQRT(9)");
            Assert.AreEqual(12d, result);
        }

        [TestMethod]
        public void ShouldRetrieveSingleValueFromEPPlus()
        {
            var result = _parser.Parse("=INT(A2)");
            Assert.AreEqual(2, result);
        }

        [TestMethod, Ignore]
        public void CircularRefsTest()
        {
            var ws = _package.Workbook.Worksheets.First();
            ws.Cells["AK5"].Value = "=SUM(AM4:AM6,AO6)";
            ws.Cells["AM4"].Value = "4";
            ws.Cells["AM5"].Value = "=SUM(AM4)";
            ws.Cells["AM6"].Value = "2";
            ws.Cells["AO6"].Value = "=SUM(AM4:AM6)";

            var result = _parser.Parse("=Int(AK5)");
            Assert.AreEqual(20, result);
        }

        [TestMethod]
        public void ShouldReturnExcelCells()
        {
            var ws = _package.Workbook.Worksheets.First();
            ws.Cells["C1"].Value = 3;
            ws.Cells["C2"].Value = 4;
            ws.Cells["D1"].Value = 2;
            ws.Cells["D2"].Value = 2;
            ws.Cells["E1"].Value = "=SUM(C1:D2)";

            var result = _parser.Parse("=Int(E1)");
            Assert.AreEqual(11, result);
        }
    }
}
