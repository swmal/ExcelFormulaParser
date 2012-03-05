using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Tests.LexicalAnalysis
{
    [TestClass]
    public class TokenFactoryTests
    {
        private ITokenFactory _tokenFactory;

        [TestInitialize]
        public void Setup()
        {
            _tokenFactory = new TokenFactory();
        }

        [TestMethod]
        public void ShouldCreateAStringToken()
        {
            var input = "'";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("'", token.Value);
            Assert.AreEqual(TokenType.String, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreatePlusAsOperatorToken()
        {
            var input = "+";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("+", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateMinusAsOperatorToken()
        {
            var input = "-";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("-", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateMultiplyAsOperatorToken()
        {
            var input = "*";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("*", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateDivideAsOperatorToken()
        {
            var input = "/";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("/", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateIntegerAsIntegerToken()
        {
            var input = "23";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("23", token.Value);
            Assert.AreEqual(TokenType.Integer, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateDecimalAsDecimalToken()
        {
            var input = "23.3";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("23.3", token.Value);
            Assert.AreEqual(TokenType.Decimal, token.TokenType);
        }
    }
}
