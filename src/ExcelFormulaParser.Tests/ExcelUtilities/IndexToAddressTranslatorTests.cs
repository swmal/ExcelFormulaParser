using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests.ExcelUtilities
{
    [TestClass]
    public class IndexToAddressTranslatorTests
    {
        private ExcelDataProvider _excelDataProvider;
        private IndexToAddressTranslator _indexToAddressTranslator;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _indexToAddressTranslator = new IndexToAddressTranslator(_excelDataProvider);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ShouldThrowIfExcelDataProviderIsNull()
        {
            new IndexToAddressTranslator(null);
        }

        [TestMethod]
        public void ShouldTranslate0And0ToA1()
        {
            var result = _indexToAddressTranslator.ToAddress(0, 0);
            Assert.AreEqual("A1", result);
        }

        [TestMethod]
        public void ShouldTranslate26And0ToAA1()
        {
            var result = _indexToAddressTranslator.ToAddress(26, 0);
            Assert.AreEqual("AA1", result);
        }

        [TestMethod]
        public void ShouldTranslate26x26plus25And0ToZZ1()
        {
            var result = _indexToAddressTranslator.ToAddress(26*26+25, 0);
            Assert.AreEqual("ZZ1", result);
        }

        [TestMethod]
        public void ShouldTranslate26x26plus26And4ToAAA5()
        {
            var result = _indexToAddressTranslator.ToAddress(26 * 26 + 26, 4);
            Assert.AreEqual("AAA5", result);
        }
    }
}
