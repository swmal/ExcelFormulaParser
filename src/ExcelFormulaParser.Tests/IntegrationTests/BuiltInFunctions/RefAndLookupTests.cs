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
    public class RefAndLookupTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_excelDataProvider);
        }

        [TestMethod]
        public void VLookupShouldReturnAResult()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(0, 1)).Return(new ExcelCell(1, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(1, 0)).Return(new ExcelCell(2, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(1, 1)).Return(new ExcelCell(5, null, 0, 0));
            var result = _parser.Parse("VLOOKUP(2, " + lookupAddress + ", 1)");
            Assert.AreEqual(5, result);
        }
    }
}
