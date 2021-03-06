﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.Excel;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.LexicalAnalysis
{
    [TestClass]
    public class NegationTests
    {
        private SourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            var context = ParsingContext.Create();
            _tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void ShouldSetNegatorOnFirstTokenIfFirstCharIsMinus()
        {
            var input = "-1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(2, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.First().TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenIfPreviousTokenIsAnOperator()
        {
            var input = "1 + -1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(4, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(2).TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInsideParenthethis()
        {
            var input = "1 + (-1 * 2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(3).TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInsideFunctionCall()
        {
            var input = "Ceiling(-1, -0.1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(2).TokenType);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(5).TokenType, "Negator after comma was not identified");
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInEnumerable()
        {
            var input = "{-1}";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(1).TokenType);
        }
    }
}
