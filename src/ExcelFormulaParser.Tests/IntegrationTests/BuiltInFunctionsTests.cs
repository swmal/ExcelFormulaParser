using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelFormulaParser.Tests.IntegrationTests
{
    [TestClass]
    public class BuiltInFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            _parser = new FormulaParser();
        }

        [TestMethod]
        public void CstrShouldReturnString()
        {
            var result = _parser.Parse("Cstr(2)");
            Assert.AreEqual("2", result);
        }

        [TestMethod]
        public void CstrShouldConcatenateWithNextExpression()
        {
            var result = _parser.Parse("Cstr(2) & 'a'");
            Assert.AreEqual("2a", result);
        }

        [TestMethod]
        public void LenShouldAddLengthUsingSuppliedOperator()
        {
            var result = _parser.Parse("Len('abc') + 2");
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void LowerShouldReturnALowerCaseString()
        {
            var result = _parser.Parse("Lower('ABC')");
            Assert.AreEqual("abc", result);
        }

        [TestMethod]
        public void UpperShouldReturnAnUpperCaseString()
        {
            var result = _parser.Parse("Upper('abc')");
            Assert.AreEqual("ABC", result);
        }
    }
}
