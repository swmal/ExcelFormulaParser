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

        [TestMethod]
        public void TodayShouldReturnAResult()
        {
            var result = _parser.Parse("Today()");
            Assert.IsInstanceOfType(DateTime.FromOADate((double)result), typeof(DateTime));
        }

        [TestMethod]
        public void NowShouldReturnAResult()
        {
            var result = _parser.Parse("now()");
            Assert.IsInstanceOfType(DateTime.FromOADate((double)result), typeof(DateTime));
        }

        [TestMethod]
        public void DayShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Day(Date(2012, 4, 2))");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void MonthShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Month(Date(2012, 4, 2))");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Year(Date(2012, 2, 2))");
            Assert.AreEqual(2012, result);
        }
    }
}
