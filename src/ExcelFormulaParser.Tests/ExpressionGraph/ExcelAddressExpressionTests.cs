using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExpressionGraph;
using Rhino.Mocks;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.ExpressionGraph
{
    [TestClass]
    public class ExcelAddressExpressionTests
    {
        private ParsingContext _parsingContext;
        private ParsingScope _scope;

        [TestInitialize]
        public void Setup()
        {
            _parsingContext = ParsingContext.Create();
            _scope = _parsingContext.Scopes.NewScope();
        }

        [TestCleanup]
        public void Cleanup()
        {
            _scope.Dispose();
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfExcelDataProviderIsNull()
        {
            new ExcelAddressExpression("A1", null, _parsingContext);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfParsingContextIsNull()
        {
            new ExcelAddressExpression("A1", MockRepository.GenerateStub<ExcelDataProvider>(), null);
        }

        [TestMethod]
        public void ShouldCallReturnResultFromProvider()
        {
            var expectedAddress = "A1";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddress))
                .Return(new object[] { 1 });

            var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
            var result = expression.Compile();
            Assert.AreEqual(1, result.Result);
        }
    }
}
