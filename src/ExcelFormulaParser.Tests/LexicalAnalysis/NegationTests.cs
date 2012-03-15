using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.VBA;
using ExcelFormulaParser.Engine.VBA.Functions;

namespace ExcelFormulaParser.Tests.LexicalAnalysis
{
    [TestClass]
    public class NegationTests
    {
        private SourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            _tokenizer = new SourceCodeTokenizer();
            FunctionRepository.LoadModule(new BuiltInFunctions());
        }

        [TestCleanup]
        public void Cleanup()
        {
            FunctionRepository.Clear();
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
    }
}
