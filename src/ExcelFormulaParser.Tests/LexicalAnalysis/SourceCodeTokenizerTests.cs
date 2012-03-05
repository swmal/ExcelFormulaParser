using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.LexicalAnalysis;

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
    }
}
