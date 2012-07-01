using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.Logical;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class LogicalFunctionsTests
    {
        private ParsingContext _parsingContext;

        [TestMethod]
        public void IfShouldReturnCorrectResult()
        {
            var func = new If();
            var args = new object[] { true, "A", "B" };
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual("A", result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIsTrue()
        {
            var func = new Not();
            var args = new object[] { true };
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnTrueIfArgumentIs0()
        {
            var func = new Not();
            var args = new object[] { 0 };
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIs1()
        {
            var func = new Not();
            var args = new object[] { 1 };
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnTrueIfAllArgumentsAreTrue()
        {
            var func = new And();
            var args = new object[] { true, true, true };
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnTrueIfAllArgumentsAreTrueOr1()
        {
            var func = new And();
            var args = new object[] { true, true, 1, true, 1 };
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnFalseIfOneArgumentIsFalse()
        {
            var func = new And();
            var args = new object[] { true, false, true };
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnFalseIfOneArgumentIs0()
        {
            var func = new And();
            var args = new object[] { true, 0, true };
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void OrShouldReturnTrueIfOneArgumentIsTrue()
        {
            var func = new Or();
            var args = new object[] { true, false, false };
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }
    }
}
