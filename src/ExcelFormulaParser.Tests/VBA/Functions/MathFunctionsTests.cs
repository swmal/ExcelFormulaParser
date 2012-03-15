using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.Math;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class MathFunctionsTests
    {
        [TestMethod]
        public void PiShouldReturnPIConstant()
        {
            var expectedValue = (decimal)Math.Round(Math.PI, 14);
            var func = new Pi();
            var args = new object[0];
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0()
        {
            var expectedValue = 22.36m;
            var func = new Ceiling();
            var args = new object[]{22.35m, 0.01};
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
        {
            var expectedValue = -22.4m;
            var func = new Ceiling();
            var args = new object[] { -22.35m, -0.1 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1()
        {
            var expectedValue = 23m;
            var func = new Ceiling();
            var args = new object[] { 22.35m, 1 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10()
        {
            var expectedValue = 30m;
            var func = new Ceiling();
            var args = new object[] { 22.35m, 10 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative()
        {
            var expectedValue = -30m;
            var func = new Ceiling();
            var args = new object[] { -22.35m, -10 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
        {
            var expectedValue = 30m;
            var func = new Ceiling();
            var args = new object[] { 22.35m, -1 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void SumShouldCalculate2Plus3AndReturn5()
        {
            var func = new Sum();
            var args = new object[] { 2, 3 };
            var result = func.Execute(args);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumShouldCalculateEnumerableOf2Plus5Plus3AndReturn10()
        {
            var func = new Sum();
            var args = new object[] { new object[]{2, 5}, 3 };
            var result = func.Execute(args);
            Assert.AreEqual(10d, result.Result);
        }
    }
}
