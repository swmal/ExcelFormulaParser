using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class ParsingScopeTests
    {
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;
        private ParsingScopes _parsingScopes;

        [TestInitialize]
        public void Setup()
        {
            _lifeTimeEventHandler = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            _parsingScopes = MockRepository.GenerateStub<ParsingScopes>(_lifeTimeEventHandler);
        }

        [TestMethod]
        public void ConstructorShouldSetAddress()
        {
            var expectedAddress =  RangeAddress.Parse("A1");
            var scope = new ParsingScope(_parsingScopes, expectedAddress);
            Assert.AreEqual(expectedAddress, scope.Address);
        }

        [TestMethod]
        public void ConstructorShouldSetParent()
        {
            var parent = new ParsingScope(_parsingScopes, RangeAddress.Parse("A1"));
            var scope = new ParsingScope(_parsingScopes, parent, RangeAddress.Parse("A2"));
            Assert.AreEqual(parent, scope.Parent);
        }

        [TestMethod]
        public void ScopeShouldCallKillScopeOnDispose()
        {
            var scope = new ParsingScope(_parsingScopes, RangeAddress.Parse("A1"));
            ((IDisposable)scope).Dispose();
            _parsingScopes.AssertWasCalled(x => x.KillScope(scope));
        }
    }
}
