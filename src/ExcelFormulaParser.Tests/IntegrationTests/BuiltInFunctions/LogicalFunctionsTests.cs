using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class LogicalFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            _parser = new FormulaParser();
        }

        [TestMethod]
        public void IfShouldReturnCorrectResult()
        {
            var result = _parser.Parse("If(2 < 3, 1, 2)");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void IIfShouldReturnCorrectResultWhenInnerFunctionExists()
        {
            var result = _parser.Parse("If(NOT(Or(true, FALSE)), 1, 2)");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void NotShouldReturnCorrectResult()
        {
            var result = _parser.Parse("not(true)");
            Assert.IsFalse((bool)result);

            result = _parser.Parse("NOT(false)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void AndShouldReturnCorrectResult()
        {
            var result = _parser.Parse("And(true, 1)");
            Assert.IsTrue((bool)result);

            result = _parser.Parse("AND(true, true, 1, false)");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void OrShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Or(FALSE, 0)");
            Assert.IsFalse((bool)result);

            result = _parser.Parse("OR(true, true, 1, false)");
            Assert.IsTrue((bool)result);
        }
    }
}
