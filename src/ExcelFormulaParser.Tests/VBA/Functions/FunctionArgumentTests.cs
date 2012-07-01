using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions;
using ExcelFormulaParser.Engine.VBA;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class FunctionArgumentTests
    {
        [TestMethod]
        public void ShouldSetExcelState()
        {
            var arg = new FunctionArgument(2);
            arg.SetExcelStateFlag(ExcelCellState.HiddenCell);
            Assert.IsTrue(arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
        }

        [TestMethod]
        public void ExcelStateFlagIsSetShouldReturnFalseWhenNotSet()
        {
            var arg = new FunctionArgument(2);
            Assert.IsFalse(arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
        }
    }
}
