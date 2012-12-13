using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.EPPlus.Tests.Helpers;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.EPPlus.Tests
{
    [TestClass]
    public class CalculationChainTests
    {
        [TestMethod]
        public void ChainTest1()
        {
            var cbt = new ChainTestBuilder();
            using (var package = cbt.Build())
            {
                package.SaveAs(new System.IO.FileInfo("c:\\Temp\\chaintest.xlsx"));
                var provider = new EPPlusExcelDataProvider(package);
                var startTime = DateTime.Now;
                var parser = new FormulaParser(provider);
                var result = parser.ParseAt("C1");
                var elapsed = DateTime.Now.Subtract(startTime);
            }
        }
    }
}
