using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Tests.TestHelpers;

namespace ExcelFormulaParser.Tests.Excel.Functions
{
    [TestClass]
    public class ExcelFunctionTests
    {
        private class ExcelFunctionTester : ExcelFunction
        {
            public IEnumerable<double> ArgsToDoubleEnumerableImpl(IEnumerable<FunctionArgument> args)
            {
                return ArgsToDoubleEnumerable(args);
            }
            #region Other members
            public override Engine.ExpressionGraph.CompileResult Execute(IEnumerable<FunctionArgument> arguments, Engine.ParsingContext context)
            {
                throw new NotImplementedException();
            }
            #endregion
        }

        [TestMethod]
        public void ArgsToDoubleEnumerableShouldHandleInnerEnumerables()
        {
            var args = FunctionsHelper.CreateArgs(1, 2, FunctionsHelper.CreateArgs(3, 4));
            var tester = new ExcelFunctionTester();
            var result = tester.ArgsToDoubleEnumerableImpl(args);
            Assert.AreEqual(4, result.Count());
        }
    }
}
