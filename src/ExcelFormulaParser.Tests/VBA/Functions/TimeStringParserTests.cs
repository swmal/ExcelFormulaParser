using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.DateTime;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class TimeStringParserTests
    {
        private double GetSerialNumber(int hour, int minute, int second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
        }

        [TestMethod]
        public void ParseShouldIdentifyPatternAndReturnCorrectResult()
        {
            var parser = new TimeStringParser();
            var result = parser.Parse("10:12:55");
            Assert.AreEqual(GetSerialNumber(10, 12, 55), result);
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void ParseShouldThrowExceptionIfSecondIsOutOfRange()
        {
            var parser = new TimeStringParser();
            var result = parser.Parse("10:12:60");
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void ParseShouldThrowExceptionIfMinuteIsOutOfRange()
        {
            var parser = new TimeStringParser();
            var result = parser.Parse("10:60:55");
        }
    }
}
