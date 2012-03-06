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

        [TestMethod, Ignore]
        public void CstrShouldConcatenateWithNextExpression()
        {
            // TODO: in this test Parse returns "2"
            var result = _parser.Parse("Cstr(2) & 'a'");
            Assert.AreEqual("2a", result);
        }
    }
}
