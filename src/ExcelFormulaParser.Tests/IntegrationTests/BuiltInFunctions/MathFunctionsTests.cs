using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class MathFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            _parser = new FormulaParser();
        }

        [TestMethod]
        public void PowerShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Power(3, 3)");
            Assert.AreEqual(27m, result);
        }

        [TestMethod]
        public void SqrtShouldReturnCorrectResult()
        {
            var result = _parser.Parse("sqrt(9)");
            Assert.AreEqual(3m, result);
        }
    }
}
