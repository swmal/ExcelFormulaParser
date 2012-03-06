using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.VBA.Functions.Text;

namespace ExcelFormulaParser.Tests.VBA.Functions.Text
{
    [TestClass]
    public class TextFunctionsTests
    {
        [TestMethod]
        public void CStrShouldConvertNumberToString()
        {
            var func = new CStr();
            var result = func.Execute(new object[] { 1 });
            Assert.AreEqual(DataType.String, result.DataType);
            Assert.AreEqual("1", result.Result);
        }

        [TestMethod]
        public void LenShouldReturnStringsLength()
        {
            var func = new Len();
            var result = func.Execute(new object[] { "abc" });
            Assert.AreEqual(3, result.Result);
        }
    }
}
