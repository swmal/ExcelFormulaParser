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

        [TestMethod]
        public void ShouldDivideDecimalWithIntegerCorrectly()
        {
            var result = _parser.Parse("2.5/2");
            Assert.AreEqual(1.25m, result);
        }

        [TestMethod]
        public void ShouldMultiplyDecimalWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 * 1.5");
            Assert.AreEqual(3.75m, result);
        }

        [TestMethod]
        public void ThreeGreaterThanTwoShouldBeTrue()
        {
            var result = _parser.Parse("3 > 2");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanTwoShouldBeFalse()
        {
            var result = _parser.Parse("3 < 2");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 <= 3");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanOrEqualToTwoDotThreeShouldBeFalse()
        {
            var result = _parser.Parse("3 <= 2.3");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void ThreeGreaterThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 >= 3");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TwoDotTwoGreaterThanOrEqualToThreeShouldBeFalse()
        {
            var result = _parser.Parse("2.2 >= 3");
            Assert.IsFalse((bool)result);
        }
    }
}
