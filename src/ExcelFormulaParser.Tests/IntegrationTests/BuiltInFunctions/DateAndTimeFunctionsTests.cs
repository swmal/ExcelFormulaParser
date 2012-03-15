using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class DateAndTimeFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            _parser = new FormulaParser();
        }

        [TestMethod]
        public void DateShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Date(2012, 2, 2)");
            Assert.AreEqual(new DateTime(2012, 2, 2).ToOADate(), result);
        }
    }
}
