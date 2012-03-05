using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.VBA.Operators;

namespace ExcelFormulaParser.Tests.ExpressionGraph
{
    [TestClass]
    public class ExpressionConverterTests
    {
        private IExpressionConverter _converter;

        [TestInitialize]
        public void Setup()
        {
            _converter = new ExpressionConverter();
        }

        [TestMethod]
        public void ToStringExpressionShouldConvertIntegerExpressionToStringExpression()
        {
            var integerExpression = new IntegerExpression("2");
            var result = _converter.ToStringExpression(integerExpression);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual("2", result.Compile().Result);
        }

        [TestMethod]
        public void ToStringExpressionShouldCopyOperatorToStringExpression()
        {
            var integerExpression = new IntegerExpression("2");
            integerExpression.Operator = Operator.Plus;
            var result = _converter.ToStringExpression(integerExpression);
            Assert.AreEqual(integerExpression.Operator, result.Operator);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateIntegerExpressionIfCompileResultIsInteger()
        {
            var compileResult = new CompileResult(1, DataType.Integer);
            var result = _converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(IntegerExpression));
            Assert.AreEqual(1, result.Compile().Result);
        }
    }
}
