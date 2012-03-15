﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class MathFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
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

        [TestMethod]
        public void PiShouldReturnCorrectResult()
        {
            var expectedValue = (decimal)Math.Round(Math.PI, 14);
            var result = _parser.Parse("Pi()");
            Assert.AreEqual(expectedValue, result);
        }

        [TestMethod]
        public void CeilingShouldReturnCorrectResult()
        {
            var expectedValue = 22.4m;
            var result = _parser.Parse("ceiling(22.35, 0.1)");
            Assert.AreEqual(expectedValue, result);
        }

        [TestMethod]
        public void SumShouldReturnCorrectResultWithInts()
        {
            var result = _parser.Parse("sum(1, 2)");
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void SumShouldReturnCorrectResultWithDecimals()
        {
            var result = _parser.Parse("sum(1,2.5)");
            Assert.AreEqual(3.5d, result);
        }

        [TestMethod]
        public void SumShouldReturnCorrectResultWithEnumerable()
        {
            var result = _parser.Parse("sum({1,2,3}, 2.5)");
            Assert.AreEqual(8.5d, result);
        }
    }
}
