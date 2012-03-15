using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.Math;
using System.Threading;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class MathFunctionsTests
    {
        [TestMethod]
        public void PiShouldReturnPIConstant()
        {
            var expectedValue = (double)Math.Round(Math.PI, 14);
            var func = new Pi();
            var args = new object[0];
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0()
        {
            var expectedValue = 22.36d;
            var func = new Ceiling();
            var args = new object[]{22.35d, 0.01};
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
        {
            var expectedValue = -22.4d;
            var func = new Ceiling();
            var args = new object[] { -22.35d, -0.1 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1()
        {
            var expectedValue = 23d;
            var func = new Ceiling();
            var args = new object[] { 22.35d, 1 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10()
        {
            var expectedValue = 30d;
            var func = new Ceiling();
            var args = new object[] { 22.35d, 10 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative()
        {
            var expectedValue = -30d;
            var func = new Ceiling();
            var args = new object[] { -22.35d, -10 };
            var result = func.Execute(args);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
        {
            var expectedValue = 30d;
            var func = new Ceiling();
            var args = new object[] { 22.35d, -1 };
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

        [TestMethod]
        public void StdevShouldCalculateCorrectResult()
        {
            var func = new Stdev();
            var args = new object[] { 1, 3, 5 };
            var result = func.Execute(args);
            Assert.AreEqual(1.63299d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void StdevPShouldCalculateCorrectResult()
        {
            var func = new StdevP();
            var args = new object[] { 2, 3, 4 };
            var result = func.Execute(args);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void ExpShouldCalculateCorrectResult()
        {
            var func = new Exp();
            var args = new object[] { 4 };
            var result = func.Execute(args);
            Assert.AreEqual(54.59815003d, System.Math.Round((double)result.Result, 8));
        }

        [TestMethod]
        public void MaxShouldCalculateCorrectResult()
        {
            var func = new Max();
            var args = new object[] { 4, 2, 5, 2 };
            var result = func.Execute(args);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void MinShouldCalculateCorrectResult()
        {
            var func = new Min();
            var args = new object[] { 4, 2, 5, 2 };
            var result = func.Execute(args);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResult()
        {
            var expectedResult = (4d + 2d + 5d + 2d) / 4d;
            var func = new Average();
            var args = new object[] { 4d, 2d, 5d, 2d };
            var result = func.Execute(args);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResultWithEnumerableAndBoolMembers()
        {
            var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            var func = new Average();
            var args = new object[] { new object[]{ 4d, 2d }, 5d, 2d, true };
            var result = func.Execute(args);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResult()
        {
            var func = new Round();
            var args = new object[] { 2.3433, 3 };
            var result = func.Execute(args);
            Assert.AreEqual(2.343d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResultWhenNbrOfDecimalsIsNegative()
        {
            var func = new Round();
            var args = new object[] { 9333, -3 };
            var result = func.Execute(args);
            Assert.AreEqual(9000d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1()
        {
            var func = new Floor();
            var args = new object[] { 26.75d, 0.1 };
            var result = func.Execute(args);
            Assert.AreEqual(26.7d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIs1()
        {
            var func = new Floor();
            var args = new object[] { 26.75d, 1 };
            var result = func.Execute(args);
            Assert.AreEqual(26d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1()
        {
            var func = new Floor();
            var args = new object[] { -26.75d, -1 };
            var result = func.Execute(args);
            Assert.AreEqual(-26d, result.Result);
        }

        [TestMethod]
        public void RandShouldReturnAValueBetween0and1()
        {
            var func = new Rand();
            var args = new object[0];
            var result1 = func.Execute(args);
            Assert.IsTrue(((double)result1.Result) > 0 && ((double) result1.Result) < 1);
            var result2 = func.Execute(args);
            Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
            Assert.IsTrue(((double)result2.Result) > 0 && ((double)result2.Result) < 1);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
        {
            var func = new RandBetween();
            var args = new object[] { 1, 5 };
            var result = func.Execute(args);
            CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
        {
            var func = new RandBetween();
            var args = new object[] { -5, 0 };
            var result = func.Execute(args);
            CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
        }

        [TestMethod]
        public void CountShouldReturnNumberOfNumericItems()
        {
            var func = new Count();
            var args = new object[] { 1d, 2m, 3, new DateTime(2012, 4, 1), "4" };
            var result = func.Execute(args);
            Assert.AreEqual(4d, result.Result);
        }


        [TestMethod]
        public void CountShouldIncludeEnumerableMembers()
        {
            var func = new Count();
            var args = new object[] { 1d, new object[]{12, 13} };
            var result = func.Execute(args);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void CountAShouldReturnNumberOfNonWhitespaceItems()
        {
            var func = new CountA();
            var args = new object[] { 1d, 2m, 3, new DateTime(2012, 4, 1), "4", null, string.Empty };
            var result = func.Execute(args);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIncludeEnumerableMembers()
        {
            var func = new CountA();
            var args = new object[] { 1d, new object[] { 12, 13 } };
            var result = func.Execute(args);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void ProductShouldMultiplyArguments()
        {
            var func = new Product();
            var args = new object[] { 2d, 2d, 4d };
            var result = func.Execute(args);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleEnumerable()
        {
            var func = new Product();
            var args = new object[] { 2d, 2d, new object[]{ 4d, 2d } };
            var result = func.Execute(args);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleFirstItemIsEnumerable()
        {
            var func = new Product();
            var args = new object[] { new object[] { 4d, 2d }, 2d, 2d };
            var result = func.Execute(args);
            Assert.AreEqual(32d, result.Result);
        }
    }
}
