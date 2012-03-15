using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.Logical;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class LogicalFunctionsTests
    {
        [TestMethod]
        public void IfShouldReturnCorrectResult()
        {
            var func = new If();
            var args = new object[] { true, "A", "B" };
            var result = func.Execute(args);
            Assert.AreEqual("A", result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIsTrue()
        {
            var func = new Not();
            var args = new object[] { true };
            var result = func.Execute(args);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnTrueIfArgumentIs0()
        {
            var func = new Not();
            var args = new object[] { 0 };
            var result = func.Execute(args);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIs1()
        {
            var func = new Not();
            var args = new object[] { 1 };
            var result = func.Execute(args);
            Assert.IsFalse((bool)result.Result);
        }
    }
}
