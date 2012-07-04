using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class ParsingScopesTest
    {
        private ParsingScopes _parsingScopes;
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;

        [TestInitialize]
        public void Setup()
        {
            _lifeTimeEventHandler = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            _parsingScopes = new ParsingScopes(_lifeTimeEventHandler);
        }

        [TestMethod]
        public void CreatedScopeShouldBeCurrentScope()
        {
            using (var scope = _parsingScopes.NewScope())
            {
                Assert.AreEqual(_parsingScopes.Current, scope);
            }
        }

        [TestMethod]
        public void CurrentScopeShouldHandleNestedScopes()
        {
            using (var scope1 = _parsingScopes.NewScope())
            {
                Assert.AreEqual(_parsingScopes.Current, scope1);
                using (var scope2 = _parsingScopes.NewScope())
                {
                    Assert.AreEqual(_parsingScopes.Current, scope2);
                }
                Assert.AreEqual(_parsingScopes.Current, scope1);
            }
            Assert.IsNull(_parsingScopes.Current);
        }

        [TestMethod]
        public void CurrentScopeShouldBeNullWhenScopeHasTerminated()
        {
            using (var scope = _parsingScopes.NewScope())
            { }
            Assert.IsNull(_parsingScopes.Current);
        }

        [TestMethod]
        public void LifetimeEventHandlerShouldBeCalled()
        {
            using (var scope = _parsingScopes.NewScope())
            { }
            _lifeTimeEventHandler.AssertWasCalled(x => x.ParsingCompleted());
        }
    }
}