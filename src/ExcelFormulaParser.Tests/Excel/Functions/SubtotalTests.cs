using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.Excel.Functions.Math;
using ExcelFormulaParser.Tests.TestHelpers;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.Excel.Functions
{
    [TestClass]
    public class SubtotalTests
    {
        private ParsingContext _context;

        [TestMethod]
        public void ShouldCalculateAverageWhenCalcTypeIs1()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(1, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(30d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateCountWhenCalcTypeIs2()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(2, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateCountAWhenCalcTypeIs3()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(3, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateMaxWhenCalcTypeIs4()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(4, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(50d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateMinWhenCalcTypeIs5()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(5, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateProductWhenCalcTypeIs6()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(6, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(12000000d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateStdevWhenCalcTypeIs7()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(7, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            var resultRounded = Math.Round((double)result.Result, 5);
            Assert.AreEqual(15.81139d, resultRounded);
        }

        [TestMethod]
        public void ShouldCalculateStdevPWhenCalcTypeIs8()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(8, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            var resultRounded = Math.Round((double)result.Result, 8);
            Assert.AreEqual(14.14213562, resultRounded);
        }

        [TestMethod]
        public void ShouldCalculateSumPWhenCalcTypeIs9()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(9, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(150d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateVarWhenCalcTypeIs10()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(10, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(250d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateVarPWhenCalcTypeIs11()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(11, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(200d, result.Result);
        }
    }
}
