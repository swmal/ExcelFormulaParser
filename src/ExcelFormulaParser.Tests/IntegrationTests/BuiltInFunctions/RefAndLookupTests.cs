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
            var returnDict = new Dictionary<int, IList<ExcelCell>>();
            returnDict.Add(0, new List<ExcelCell>() { new ExcelCell(1, "A1", 0, 0), new ExcelCell(1, "B1", 1, 0) });
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetLookupArray(lookupAddress)).Return(returnDict);
            _excelDataProvider.Stub(x => x.GetRangeValues("A1")).Return(new List<ExcelCell> { new ExcelCell(1, null, 0, 0) });
            _excelDataProvider.Stub(x => x.GetRangeValues("B1")).Return(new List<ExcelCell> { new ExcelCell(1, null, 0, 0) });
            var result = _parser.Parse("VLOOKUP(1, " + lookupAddress + ", 2)");
        }
    }
}
