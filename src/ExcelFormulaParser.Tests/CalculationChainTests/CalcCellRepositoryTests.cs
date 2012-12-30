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
    public class CalcCellRepositoryTests
    {
        private CalcCellRepository _repository;
        private IdProvider _idProvider = new IntegerIdProvider();
        private CalcChainContext _chainContext;

        [TestInitialize]
        public void Setup()
        {
            _chainContext = CalcChainContext.Create(_idProvider);
            _repository = CalcCellRepository.Create(_chainContext);
        }

        [TestMethod]
        public void AddOrGetShouldAddACellToRepo()
        {
            _repository.AddOrGet("A1");
            var cell = _repository.GetCell("A1");
            Assert.AreEqual("A1", cell.Address);
        }

        [TestMethod]
        public void GetShouldReturnById()
        {
            var cell = _repository.AddOrGet("A1");
            var cell2 = _repository.GetCell(cell.Id);
            Assert.AreEqual(cell, cell2);
        }

        [TestMethod]
        public void AddOrGetShouldReturnTheSameInstance()
        {
            var cell = _repository.AddOrGet("A1");
            var cell2 = _repository.AddOrGet("A1");
            Assert.AreEqual(cell, cell2);
        }
    }
}
