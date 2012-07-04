using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExpGraph = ExcelFormulaParser.Engine.ExpressionGraph.ExpressionGraph;
using ExcelFormulaParser.Engine.Excel.Operators;

namespace ExcelFormulaParser.Tests.ExpressionGraph
{
    [TestClass]
    public class ExpressionCompilerTests
    {
        private IExpressionCompiler _expressionCompiler;
        private ExpGraph _graph;
        
        [TestInitialize]
        public void Setup()
        {
            _expressionCompiler = new ExpressionCompiler();
            _graph = new ExpGraph();
        }

        [TestMethod]
        public void ShouldCompileTwoInterExpressionsToCorrectResult()
        {
            var exp1 = new IntegerExpression("2");
            exp1.Operator = Operator.Plus;
            _graph.Add(exp1);
            var exp2 = new IntegerExpression("2");
            _graph.Add(exp2);

            var result = _expressionCompiler.Compile(_graph.Expressions);

            Assert.AreEqual(4, result.Result);
        }


        [TestMethod]
        public void CompileShouldMultiplyGroupExpressionWithFollowingIntegerExpression()
        {
            var groupExpression = new GroupExpression();
            groupExpression.AddChild(new IntegerExpression("2"));
            groupExpression.Children.First().Operator = Operator.Plus;
            groupExpression.AddChild(new IntegerExpression("3"));
            groupExpression.Operator = Operator.Multiply;

            _graph.Add(groupExpression);
            _graph.Add(new IntegerExpression("2"));

            var result = _expressionCompiler.Compile(_graph.Expressions);

            Assert.AreEqual(10, result.Result);
        }

        [TestMethod]
        public void CompileShouldCalculateMultipleExpressionsAccordingToPrecedence()
        {
            var exp1 = new IntegerExpression("2");
            exp1.Operator = Operator.Multiply;
            _graph.Add(exp1);
            var exp2 = new IntegerExpression("2");
            exp2.Operator = Operator.Plus;
            _graph.Add(exp2);
            var exp3 = new IntegerExpression("2");
            exp3.Operator = Operator.Multiply;
            _graph.Add(exp3);
            var exp4 = new IntegerExpression("2");
            _graph.Add(exp4);

            var result = _expressionCompiler.Compile(_graph.Expressions);

            Assert.AreEqual(8, result.Result);
        }
    }
}
