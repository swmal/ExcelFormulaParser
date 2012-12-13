using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Engine.Exceptions;
using ExcelFormulaParser.EPPlus.Tests.Helpers;

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
            var result = _parser.Parse("SUM(A1:A5) - SQRT(9)");
            Assert.AreEqual(12d, result);
        }

        [TestMethod]
        public void ShouldRetrieveSingleValueFromEPPlus()
        {
            var result = _parser.Parse("INT(A2)");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void CircularRefsTest2()
        {
            var ws = _package.Workbook.Worksheets.First();
            ws.Cells["AK5"].Formula = "SUM(AM4:AM6,AO6)";
            ws.Cells["AM4"].Value = 4;
            ws.Cells["AM5"].Formula = "SUM(AM4)";
            ws.Cells["AM6"].Value = 2;
            ws.Cells["AO6"].Formula = "SUM(AM4:AM6)";

            var result = _parser.ParseAt("AK5");
            Assert.AreEqual(20d, result);
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void ShouldDetectCircularRef()
        {
            var ws = _package.Workbook.Worksheets.First();
            ws.Cells["A1"].Formula = "SUM(A2)";
            ws.Cells["A2"].Formula = "SUM(A1)";
            _parser.ParseAt("A1");
        }

        [TestMethod]
        public void ShouldHandleMatrixValues()
        {
            var ws = _package.Workbook.Worksheets.First();
            ws.Cells["C1"].Value = 3;
            ws.Cells["C2"].Value = 4;
            ws.Cells["D1"].Value = 2;
            ws.Cells["D2"].Value = 2;
            ws.Cells["E1"].Formula = "SUM(B1:D2)";

            var result = _parser.ParseAt("E1");
            Assert.AreEqual(11d, result);
        }

        //[TestMethod]
        //public void TestParser()
        //{
        //    using (var package = new ExcelPackage(new FileInfo(@"c:\temp\Test.xlsx")))
        //    {
        //        var provider = new EPPlusExcelDataProvider(package);
        //        var parser = new FormulaParser(provider);
        //        var result = parser.ParseAt("B6");
        //        Assert.AreEqual(30d, result);
        //    }
        //}

        [TestMethod]
        public void PerformanceTest()
        {
            _workSheet.Cells["B1"].Formula = "SUM(A1:A4) * 2 - AVERAGE(A1:A3)";
            var startTime = DateTime.Now;
            int nIterations = 0;
            while(true)
            {
                 _parser.ParseAt("B1");
                //_parser.Parse("SUM({1,2,3,4}) * 2 - AVERAGE({1,2,3})");
                nIterations++;
                if (DateTime.Now.Subtract(startTime).TotalMilliseconds > 1000)
                    break;
            }
            Console.WriteLine("Result: " + nIterations);
            // Current result on Parse("B1"): 5800
            // Non excel address formula: 8500
        }

        [TestMethod]
        public void DefinedNameTest()
        {
            //lopment\ExcelFormulaParser\src\ExcelFormulaParser.EPPlus.Tests\Files\lite
            var fileInfo = new FileInfo("..\\..\\..\\ExcelFormulaParser.EPPlus.Tests\\Files\\lite_olika_namn.xlsx");
            using (var package = new ExcelPackage(fileInfo))
            {
                var provider = new EPPlusExcelDataProvider(package);
                var parser = new FormulaParser(provider);
                var result = parser.ParseAt("C4");
                Assert.AreEqual(2d, result);
            }
        }

        [TestMethod, Timeout(TestTimeout.Infinite)]
        public void LoadTestFromJan()
        {
            var fileInfo = new FileInfo("c:\\temp\\xl\\Kedjor-prestanda test.xlsx");
            using(var package = new ExcelPackage(fileInfo))
            {
                var provider = new EPPlusExcelDataProvider(package);
                var parser = new FormulaParser(provider);
                var startTime = DateTime.Now;
                var result = parser.ParseAt("B1");
                var elapsed = DateTime.Now.Subtract(startTime);
            }

        }

        
    }
}
