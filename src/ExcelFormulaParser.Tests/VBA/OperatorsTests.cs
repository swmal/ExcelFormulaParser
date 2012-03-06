using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Operators;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Tests.VBA
{
    [TestClass]
    public class OperatorsTests
    {
        [TestMethod]
        public void OperatorConcatShouldConcatTwoStrings()
        {
            var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String));
            Assert.AreEqual("ab", result);
        }

        [TestMethod]
        public void OperatorConcatShouldConcatANumberAndAString()
        {
            var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String));
            Assert.AreEqual("12b", result);
        }
    }
}
