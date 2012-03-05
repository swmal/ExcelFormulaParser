using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.VBA.Operators;

namespace ExcelFormulaParser.Tests.ExpressionGraph
{
    [TestClass]
    public class ExpressionGraphBuilderTests
    {
        private IExpressionGraphBuilder _graphBuilder;

        [TestInitialize]
        public void Setup()
        {
            _graphBuilder = new ExpressionGraphBuilder();    
        }

        [TestMethod]
        public void BuildShouldNotUseStringIdentifyersWhenBuildingStringExpression()
        {
            var tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
        }

        [TestMethod]
        public void BuildShouldSetOperatorOnGroupExpressionCorrectly()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(Operator.Multiply.Operator, result.Expressions.First().Operator.Operator);

        }

        [TestMethod]
        public void BuildShouldSetChildrenOnGroupExpression()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.IsInstanceOfType(result.Expressions.First(), typeof(GroupExpression));
            Assert.AreEqual(2, result.Expressions.First().Children.Count());
        }

        [TestMethod]
        public void BuildShouldSetNextOnGroupedExpression()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.IsNotNull(result.Expressions.First().Next);
            Assert.IsInstanceOfType(result.Expressions.First().Next, typeof(IntegerExpression));

        }
    }
}
