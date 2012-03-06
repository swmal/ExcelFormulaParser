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
    public class SourceCodeTokenizerTests
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
        public void ShouldCreateTokensForStringCorrectly()
        {
            var input = "'abc123'";
            var tokens = _tokenizer.Tokenize(input);
            
            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.String, tokens.First().TokenType);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.String, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldCreateTokensForStringCorrectly2()
        {
            var input = "\"abc123\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.String, tokens.First().TokenType);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.String, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldIgnoreTokenSeparatorsInAString()
        {
            var input = "'ab(c)d'";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
        }

        [TestMethod]
        public void ShouldCreateTokensForFunctionCorrectly()
        {
            var input = "CStr(2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(4, tokens.Count());
            Assert.AreEqual(TokenType.Function, tokens.First().TokenType);
            Assert.AreEqual(TokenType.OpeningBracket, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
            Assert.AreEqual("2", tokens.ElementAt(2).Value);
            Assert.AreEqual(TokenType.ClosingBracket, tokens.Last().TokenType);
        }
    }
}
