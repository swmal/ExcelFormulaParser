using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelFormulaParser.Tests.IntegrationTests
{
    [TestClass]
    public class PrecedenceTests : FormulaParserTestBase
    {

        [TestInitialize]
        public void Setup()
        {
            _parser = new FormulaParser();
        }

        [TestMethod]
        public void ShouldCaluclateUsingPrecedenceMultiplyBeforeAdd()
        {
            var result = _parser.Parse("4 + 6 * 2");
            Assert.AreEqual(16, result);
        }

        [TestMethod]
        public void ShouldCaluclateUsingPrecedenceDivideBeforeAdd()
        {
            var result = _parser.Parse("4 + 6 / 2");
            Assert.AreEqual(7, result);
        }

        [TestMethod]
        public void ShouldCalculateTwoGroupsUsingDivideAndMultiplyBeforeSubtract()
        {
            var result = _parser.Parse("4/2 + 3 * 3");
            Assert.AreEqual(11, result);
        }

        [TestMethod]
        public void ShouldCalculateExpressionWithinParenthesisBeforeMultiply()
        {
            var result = _parser.Parse("(2+4) * 2");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void ShouldConcatAfterAdd()
        {
            var result = _parser.Parse("2 + 4 & 'abc'");
            Assert.AreEqual("6abc", result);
        }
    }
}
