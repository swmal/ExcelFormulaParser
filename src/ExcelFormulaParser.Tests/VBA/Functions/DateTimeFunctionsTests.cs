using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.DateTime;
using ExcelFormulaParser.Engine.ExpressionGraph;
using System.Threading;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class DateTimeFunctionsTests
    {
        [TestMethod]
        public void DateFunctionShouldReturnADate()
        {
            var func = new Date();
            var args = new object[] { 2012, 4, 3 };
            var result = func.Execute(args);
            Assert.AreEqual(DataType.Date, result.DataType);
        }

        [TestMethod]
        public void DateFunctionShouldReturnACorrectDate()
        {
            var expectedDate = new DateTime(2012, 4, 3);
            var func = new Date();
            var args = new object[] { 2012, 4, 3 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedDate.ToOADate(), result.Result);
        }

        [TestMethod]
        public void DateFunctionShouldMonthFromPrevYearIfMonthIsNegative()
        {
            var expectedDate = new DateTime(2011, 11, 3);
            var func = new Date();
            var args = new object[] { 2012, -1, 3 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedDate.ToOADate(), result.Result);
        }

        [TestMethod]
        public void NowFunctionShouldReturnNow()
        {
            var startTime = DateTime.Now;
            Thread.Sleep(1);
            var func = new Now();
            var args = new object[0];
            var result = func.Execute(args);
            Thread.Sleep(1);
            var endTime = DateTime.Now;
            var resultDate = DateTime.FromOADate((double)result.Result);
            Assert.IsTrue(resultDate > startTime && resultDate < endTime);
        }

        [TestMethod]
        public void TodayFunctionShouldReturnTodaysDate()
        {
            var func = new Today();
            var args = new object[0];
            var result = func.Execute(args);
            var resultDate = DateTime.FromOADate((double)result.Result);
            Assert.AreEqual(DateTime.Now.Date, resultDate);
        }

        [TestMethod]
        public void DayShouldReturnDayInMonth()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Day();
            var result = func.Execute(new object[] { date.ToOADate() });
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void MonthShouldReturnMonthOfYear()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Month();
            var result = func.Execute(new object[] { date.ToOADate() });
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectYear()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Year();
            var result = func.Execute(new object[] { date.ToOADate() });
            Assert.AreEqual(2012, result.Result);
        }
    }
}
