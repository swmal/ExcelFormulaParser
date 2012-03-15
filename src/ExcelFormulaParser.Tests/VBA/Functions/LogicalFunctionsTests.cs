using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions.Logical;

namespace ExcelFormulaParser.Tests.VBA.Functions
{
    [TestClass]
    public class LogicalFunctionsTests
    {
        [TestMethod]
        public void IIFShouldReturnCorrectResult()
        {
            var func = new IIf();
            var args = new object[] { true, "A", "B" };
            var result = func.Execute(args);
            Assert.AreEqual("A", result.Result);
        }
    }
}
