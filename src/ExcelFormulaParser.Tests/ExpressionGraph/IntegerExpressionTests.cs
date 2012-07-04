using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.Excel.Operators;

namespace ExcelFormulaParser.Tests.ExpressionGraph
{
    [TestClass]
    public class IntegerExpressionTests
    {
        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldCalulateSumCorrectly()
        {
            var exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            var result = exp1.MergeWithNext();

            Assert.AreEqual(3, result.Compile().Result);
        }

        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldSetNextPointer()
        {
            var exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            var result = exp1.MergeWithNext();

            Assert.IsNull(result.Next);
        }
    }
}
