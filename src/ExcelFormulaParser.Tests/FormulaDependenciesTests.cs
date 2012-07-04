using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class FormulaDependenciesTests
    {
        private IParsingLifetimeEventHandler _lifetimeEventHandler;
        private ParsingScopes _scopes;
        private FormulaDependencyFactory _formulaDependencyFactory;
        private FormulaDependencies _formulaDependencies;
        private RangeAddressFactory _factory;

        [TestInitialize]
        public void Setup()
        {
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
            _lifetimeEventHandler = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            _scopes = MockRepository.GenerateStub<ParsingScopes>(_lifetimeEventHandler);
            _formulaDependencyFactory = MockRepository.GenerateStub<FormulaDependencyFactory>();
            _formulaDependencies = new FormulaDependencies(_formulaDependencyFactory);
        }

        [TestMethod]
        public void DependenciesShouldHaveOneItemAfterAdd()
        {
            _formulaDependencies.AddFormulaScope(new ParsingScope(_scopes, _factory.Create("A1")));
            Assert.AreEqual(1, _formulaDependencies.Dependencies.Count());
        }

        [TestMethod]
        public void ShouldAddReferenceToParent()
        {
            var parentScope = new ParsingScope(_scopes, _factory.Create("A1"));
            var parentDependency = MockRepository.GenerateStub<FormulaDependency>(parentScope);
            _formulaDependencyFactory.Stub(x => x.Create(parentScope)).Return(parentDependency);
            _formulaDependencies.AddFormulaScope(parentScope);
            
            var childScope = new ParsingScope(_scopes, parentScope, _factory.Create("A2"));
            var childDependency = MockRepository.GenerateStub<FormulaDependency>(childScope);
            _formulaDependencyFactory.Stub(x => x.Create(childScope)).Return(childDependency);
            _formulaDependencies.AddFormulaScope(childScope);

            parentDependency.AssertWasCalled(x => x.AddReferenceTo(childScope.Address));
        }

        [TestMethod]
        public void ClearShouldClearDictionary()
        {
            var scope = new ParsingScope(_scopes, _factory.Create("A1"));
            _formulaDependencies.AddFormulaScope(scope);
            Assert.AreEqual(1, _formulaDependencies.Dependencies.Count());
            _formulaDependencies.Clear();
            Assert.AreEqual(0, _formulaDependencies.Dependencies.Count());
        }

    }
}
