using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Operators;

namespace ExcelFormulaParser.Tests.VBA
{
    [TestClass]
    public class OperatorsTests
    {
        [TestMethod]
        public void OperatorConcatShouldConcatTwoStrings()
        {
            var result = Operator.Concat.Apply("a", "b");
            Assert.AreEqual("ab", result);
        }

        [TestMethod]
        public void OperatorConcatShouldConcatANumberAndAString()
        {
            var result = Operator.Concat.Apply(12, "b");
            Assert.AreEqual("12b", result);
        }
    }
}
