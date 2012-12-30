using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.CalculationChain;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Tests.CalculationChainTests
{
    [TestClass]
    public class CalcChainTests
    {
        private CalcChain _chain;
        private CalcChainContext _chainContext;

        [TestInitialize]
        public void Setup()
        {
            _chainContext = CalcChainContext.Create(new IntegerIdProvider());
            _chain = CalcChain.Create(_chainContext.IdProvider);
        }

        [TestMethod]
        public void CountShouldIncreaseWhenACellIsAdded()
        {
            var newCell = CalcCell.Create("A1", _chainContext);
            _chain.Add(newCell);
            Assert.AreEqual(1L, _chain.Count);
        }

        [TestMethod]
        public void AddShouldSetCurrentToTheAddedCell()
        {
            var newCell = CalcCell.Create("A1", _chainContext);
            _chain.Add(newCell);
            Assert.AreEqual(newCell, _chain.Current);
        }
    }
}
