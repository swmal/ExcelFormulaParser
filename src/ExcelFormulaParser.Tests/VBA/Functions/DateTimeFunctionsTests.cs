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
        public void TodayFunctionShouldReturnToday()
        {
            var startTime = DateTime.Now;
            Thread.Sleep(1);
            var func = new Today();
            var args = new object[0];
            var result = func.Execute(args);
            Thread.Sleep(1);
            var endTime = DateTime.Now;
            var resultDate = DateTime.FromOADate((double)result.Result);
            Assert.IsTrue(resultDate > startTime && resultDate < endTime);
        }
    }
}
