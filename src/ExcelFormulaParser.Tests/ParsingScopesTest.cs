using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class ParsingScopesTest
    {
        private ParsingScopes _parsingScopes;

        [TestInitialize]
        public void Setup()
        {
            _parsingScopes = new ParsingScopes();
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
        public void CurrentScopeShouldBeNullWhenScopeHasTerminated()
        {
            using (var scope = _parsingScopes.NewScope())
            { }
            Assert.IsNull(_parsingScopes.Current);
        }
    }
}