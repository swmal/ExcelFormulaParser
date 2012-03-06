using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelFormulaParser.Tests.IntegrationTests
{
    [TestClass]
    public class BasicCalcTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            _parser = new FormulaParser();
        }

        [TestMethod]
        public void ShouldAddIntegersCorrectly()
        {
            var result = _parser.Parse("1 + 2");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void ShouldSubtractIntegersCorrectly()
        {
            var result = _parser.Parse("2 - 1");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void ShouldMultiplyIntegersCorrectly()
        {
            var result = _parser.Parse("2 * 3");
            Assert.AreEqual(6, result);
        }

        [TestMethod]
        public void ShouldDivideIntegersCorrectly()
        {
            var result = _parser.Parse("8 / 4");
            Assert.AreEqual(2, result);
        }

        [TestMethod, Ignore]
        public void ShouldDivideDecimalWithIntegerCorrectly()
        {
            // TODO: handle current casting to in the operator
            var result = _parser.Parse("2.5/2");
            Assert.AreEqual(1.25m, result);
        }
    }
}
