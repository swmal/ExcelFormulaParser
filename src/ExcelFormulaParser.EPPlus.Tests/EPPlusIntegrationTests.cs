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
    }
}
