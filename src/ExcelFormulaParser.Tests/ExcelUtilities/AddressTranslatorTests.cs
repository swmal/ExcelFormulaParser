using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Tests.ExcelUtilities
{
    [TestClass]
    public class AddressTranslatorTests
    {
        [TestMethod]
        public void ShouldTranslateRowNumber()
        {
            var translator = new AddressTranslator();
            int col, row;
            translator.ToColAndRow("A2", out col, out row);
            Assert.AreEqual(1, row);
        }

        [TestMethod]
        public void ShouldTranslateLettersToColumnIndex()
        {
            var translator = new AddressTranslator();
            int col, row;
            translator.ToColAndRow("C1", out col, out row);
            Assert.AreEqual(2, col);
            translator.ToColAndRow("AA2", out col, out row);
            Assert.AreEqual(26, col);
            translator.ToColAndRow("BC1", out col, out row);
            Assert.AreEqual(54, col);
        }
    }
}
