﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Engine.ExpressionGraph;

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

        [TestMethod]
        public void TimeShouldReturnCorrectResult()
        {
            var expectedResult = ((double)(12 * 60 * 60 + 13 * 60 + 14))/((double)(24 * 60 * 60));
            var result = _parser.Parse("Time(12, 13, 14)");
            Assert.AreEqual(expectedResult, result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResult()
        {
            var result = _parser.Parse("HOUR(Time(12, 13, 14))");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResult()
        {
            var result = _parser.Parse("minute(Time(12, 13, 14))");
            Assert.AreEqual(13, result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Second(Time(12, 13, 59))");
            Assert.AreEqual(59, result);
        }
    }
}
