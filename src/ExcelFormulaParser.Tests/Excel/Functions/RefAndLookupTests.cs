using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Tests.TestHelpers;

namespace ExcelFormulaParser.Tests.Excel.Functions
{
    [TestClass]
    public class RefAndLookupTests
    {
        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow()
        {
            var dict = new Dictionary<int, IList<ExcelCell>>();
            dict[0] = new List<ExcelCell>();
            dict[0].Add(new ExcelCell(2, null, 0, 0));
            dict[0].Add(new ExcelCell(3, null, 1, 0));

            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs(2, dict, 1);
            var result = func.Execute(args, ParsingContext.Create());
            Assert.AreEqual(3, result.Result);
        }
    }
}
