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
    public class CalcCellTests
    {
        private CalcChainContext _chainContext;

        [TestInitialize]
        public void Setup()
        {
            _chainContext = CalcChainContext.Create(new IntegerIdProvider());
        }

        [TestMethod]
        public void FactoryMethodShouldSetAddress()
        {
            var cell = CalcCell.Create("A1", _chainContext);
            Assert.AreEqual("A1", cell.Address);
        }

        [TestMethod]
        public void AddRelationToShouldAddLastToRelationsTo()
        {
            var other = CalcCell.Create("A1", _chainContext);
            var cell = CalcCell.Create("B1", _chainContext);
            cell.AddRelationTo(other, CalcChain.Create(_chainContext.IdProvider));
            Assert.AreEqual(other.Id, cell.RelationsTo.Last().CalcCellId);
        }

        [TestMethod]
        public void AddRelationToShouldNotAddDuplicates()
        {
            var other = CalcCell.Create("A1", _chainContext);
            var cell = CalcCell.Create("B1", _chainContext);
            cell.AddRelationTo(other, CalcChain.Create(_chainContext.IdProvider));
            cell.AddRelationTo(other, CalcChain.Create(_chainContext.IdProvider));
            Assert.AreEqual(1, cell.RelationsTo.Count());
        }

        [TestMethod]
        public void AddRelationFromShouldNotAddDuplicates()
        {
            var other = CalcCell.Create("A1", _chainContext);
            var cell = CalcCell.Create("B1", _chainContext);
            cell.AddRelationFrom(other, CalcChain.Create(_chainContext.IdProvider));
            cell.AddRelationFrom(other, CalcChain.Create(_chainContext.IdProvider));
            Assert.AreEqual(1, cell.RelationsFrom.Count());
        }


        [TestMethod]
        public void GetCalcChainShouldReturnNullIfNotRelationsExists()
        {
            var cell = CalcCell.Create("B1", _chainContext);
            var chainId = cell.GetCalcChainId();
            Assert.IsNull(chainId);
        }

        [TestMethod]
        public void GetCalcChainShouldReturnAChainFromExistingRelation()
        {
            var chain = CalcChain.Create(_chainContext.IdProvider);
            var cell1 = CalcCell.Create("A1", _chainContext);
            var cell2 = CalcCell.Create("A2", _chainContext);
            cell1.AddRelationTo(cell2, chain);
            cell2.AddRelationFrom(cell1, chain);
            chain.Add(cell1);
            var resultChainId = cell2.GetCalcChainId();
            Assert.AreEqual(chain.Id, resultChainId);
        }

        [TestMethod]
        public void GetCalcChainShouldReturnAChainFromLargestChainRelation()
        {
            var cellToTest = CalcCell.Create("A3", _chainContext);
            var chain1 = CalcChain.Create(_chainContext.IdProvider);

            var cell1 = CalcCell.Create("A5", _chainContext);
            cellToTest.AddRelationFrom(cell1, chain1);
            cell1.AddRelationTo(cellToTest, chain1);
            chain1.Add(cellToTest);

            var chain2 = CalcChain.Create(_chainContext.IdProvider);
            var cell2 = CalcCell.Create("A1", _chainContext);
            var cell3 = CalcCell.Create("A2", _chainContext);
            cell2.AddRelationTo(cell3, chain2);
            cell3.AddRelationFrom(cell2, chain2);
            chain2.Add(cell2);
            chain2.Add(cell3);
            
            cell3.AddRelationTo(cellToTest, chain2);
            cellToTest.AddRelationFrom(cell3, chain2);
            chain2.Add(cellToTest);

            var chain3 = CalcChain.Create(_chainContext.IdProvider);
            var cell5 = CalcCell.Create("B1", _chainContext);
            cellToTest.AddRelationFrom(cell3, chain3);
            cell5.AddRelationTo(cellToTest, chain3);
            chain3.Add(cellToTest);
            chain3.Add(cell5);

            _chainContext.AddCalcChain(chain1);
            _chainContext.AddCalcChain(chain2);
            _chainContext.AddCalcChain(chain3);

            var resultChainId = cellToTest.GetCalcChainId();
            Assert.AreEqual(chain2.Id, resultChainId);
        }
    }
}
