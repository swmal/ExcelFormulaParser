using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Tests.TestHelpers;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests.Excel.Functions
{
    [TestClass]
    public class RefAndLookupTests
    {
        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 1);
            var parsingContext = ParsingContext.Create();
            
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider.Stub(x => x.GetCellValue(0, 0)).Return(new ExcelCell(3, null, 0, 0));
            provider.Stub(x => x.GetCellValue(0, 1)).Return(new ExcelCell(1, null, 0, 0));
            provider.Stub(x => x.GetCellValue(1, 0)).Return(new ExcelCell(2, null, 0, 0));
            provider.Stub(x => x.GetCellValue(1, 1)).Return(new ExcelCell(5, null, 0, 0));

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(5, result.Result);
        }
    }
}
