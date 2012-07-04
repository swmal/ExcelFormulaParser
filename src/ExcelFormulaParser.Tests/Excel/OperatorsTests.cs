using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.Excel.Operators;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Tests.Excel
{
    [TestClass]
    public class OperatorsTests
    {
        [TestMethod]
        public void OperatorConcatShouldConcatTwoStrings()
        {
            var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String));
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void OperatorConcatShouldConcatANumberAndAString()
        {
            var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String));
            Assert.AreEqual("12b", result.Result);
        }

        [TestMethod]
        public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
        {
            var result = Operator.Eq.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer));
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OperatorEqShouldReturnTruefSuppliedValuesDiffer()
        {
            var result = Operator.Eq.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void OperatorExpShouldReturnCorrectResult()
        {
            var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer));
            Assert.AreEqual(8d, result.Result);
        }
    }
}
