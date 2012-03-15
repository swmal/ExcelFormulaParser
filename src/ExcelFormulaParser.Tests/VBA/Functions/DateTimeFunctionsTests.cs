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
        private double GetTime(int hour, int minute, int second)
        {
            var secInADay = DateTime.Today.AddDays(1).Subtract(DateTime.Today).TotalSeconds;
            var secondsOfExample = (double)(hour * 60 * 60 + minute * 60 + second);
            return secondsOfExample / secInADay;
        }
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

        [TestMethod]
        public void TimeShouldReturnACorrectSerialNumber()
        {
            var expectedResult = GetTime(10, 11, 12);
            var func = new Time();
            var result = func.Execute(new object[] { 10, 11, 12 });
            Assert.AreEqual(expectedResult, result.Result);  
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TimeShouldThrowExceptionIfSecondsIsOutOfRange()
        {
            var func = new Time();
            var result = func.Execute(new object[] { 10, 11, 60 });
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TimeShouldThrowExceptionIfMinuteIsOutOfRange()
        {
            var func = new Time();
            var result = func.Execute(new object[] { 10, 60, 12 });
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TimeShouldThrowExceptionIfHourIsOutOfRange()
        {
            var func = new Time();
            var result = func.Execute(new object[] { 24, 12, 12 });
        }

        [TestMethod]
        public void HourShouldReturnCorrectResult()
        {
            var func = new Hour();
            var result = func.Execute(new object[] { GetTime(9, 13, 14) });
            Assert.AreEqual(9, result.Result);

            result = func.Execute(new object[] { GetTime(23, 13, 14) });
            Assert.AreEqual(23, result.Result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResult()
        {
            var func = new Minute();
            var result = func.Execute(new object[] { GetTime(9, 14, 14) });
            Assert.AreEqual(14, result.Result);

            result = func.Execute(new object[] { GetTime(9, 55, 14) });
            Assert.AreEqual(55, result.Result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResult()
        {
            var func = new Second();
            var result = func.Execute(new object[] { GetTime(9, 14, 17) });
            Assert.AreEqual(17, result.Result);
        }
    }
}
