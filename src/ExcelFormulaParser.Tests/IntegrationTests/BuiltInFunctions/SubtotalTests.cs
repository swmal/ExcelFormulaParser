using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class SubtotalTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_excelDataProvider);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren()
        {
            _excelDataProvider
                .Stub(x => x.GetRangeValues("A1"))
                .Return(new List<ExcelDataItem> { new ExcelDataItem(null, "SUBTOTAL(9, A2:A3)", 0, 0)});
            _excelDataProvider
                .Stub(x => x.GetRangeValues("A2:A3"))
                .Return(new List<ExcelDataItem> { new ExcelDataItem(null, "SUBTOTAL(9, A5:A6)", 0, 1), new ExcelDataItem(2d, null, 0, 2)});
            _excelDataProvider
                .Stub(x => x.GetRangeValues("A5:A6"))
                .Return(new List<ExcelDataItem> { new ExcelDataItem(2d, null, 0, 4), new ExcelDataItem(2d, null, 0, 5)});
            var result = _parser.ParseAt("A1");
            Assert.AreEqual(2d, result);
        }
    }
}
