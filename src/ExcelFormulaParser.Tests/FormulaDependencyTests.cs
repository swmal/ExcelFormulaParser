﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine.Exceptions;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class FormulaDependencyTests
    {
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;
        private RangeAddressFactory _factory;

        [TestInitialize]
        public void Setup()
        {
            _lifeTimeEventHandler = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
        }

        [TestMethod]
        public void ConstructorShouldSetScopeId()
        {
            var scopes = new ParsingScopes(_lifeTimeEventHandler);
            var scope = new ParsingScope(scopes, RangeAddress.Empty);
            var dependency = new FormulaDependency(scope);
            Assert.AreEqual(scope.ScopeId, dependency.ScopeId);
        }

        [TestMethod]
        public void ConstructorShouldSetAddress()
        {
            var scopes = new ParsingScopes(_lifeTimeEventHandler);
            var expectedAddress = _factory.Create("A1");
            var scope = new ParsingScope(scopes, expectedAddress);
            var dependency = new FormulaDependency(scope);
            Assert.AreEqual(expectedAddress, dependency.Address);
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void AddReferenceToShouldThrowWhenReferenceToItSelf()
        {
            var lifetimeMock = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            var scopes = new ParsingScopes(lifetimeMock);
            var scope1 = scopes.NewScope(_factory.Create("A2"));
            var scope2 = scopes.NewScope(_factory.Create("A2"));
            var formulaDependency = new FormulaDependency(scope1);
            formulaDependency.AddReferenceFrom(scope2.Address);
            formulaDependency.AddReferenceTo(scope2.Address);
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void AddReferenceFromShouldThrowWhenReferenceToItSelf()
        {
            var lifetimeMock = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            var scopes = new ParsingScopes(lifetimeMock);
            var scope1 = scopes.NewScope(_factory.Create("A2"));
            var scope2 = scopes.NewScope(_factory.Create("A2"));
            var formulaDependency = new FormulaDependency(scope1);
            formulaDependency.AddReferenceTo(scope2.Address);
            formulaDependency.AddReferenceFrom(scope2.Address);
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void ShouldThrowWhenCircularReferenceIsDetectedWhenReferenceIsAdded()
        {
            var lifetimeMock = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            var scopes = new ParsingScopes(lifetimeMock);
            var scope1 = scopes.NewScope(_factory.Create("A1"));
            var scope2 = scopes.NewScope(_factory.Create("A2"));
            var formulaDependency = new FormulaDependency(scope1);
            formulaDependency.AddReferenceFrom(scope2.Address);
            formulaDependency.AddReferenceTo(scope2.Address);
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void ShouldThrowWhenCircularReferenceIsDetectedWhenReferenceByIsAdded()
        {
            var lifetimeMock = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            var scopes = new ParsingScopes(lifetimeMock);
            var scope1 = scopes.NewScope(_factory.Create("A1"));
            var scope2 = scopes.NewScope(_factory.Create("A2"));
            var formulaDependency = new FormulaDependency(scope1);
            formulaDependency.AddReferenceTo(scope2.Address);
            formulaDependency.AddReferenceFrom(scope2.Address);
        }
    }
}
