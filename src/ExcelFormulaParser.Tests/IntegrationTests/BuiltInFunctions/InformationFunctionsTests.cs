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
    public class InformationFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestMethod]
        public void IsBlankShouldReturnCorrectValue()
        {
            var result = _parser.Parse("ISBLANK(A1)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void IsNumberShouldReturnCorrectValue()
        {
            var result = _parser.Parse("ISNUMBER(1)");
            Assert.IsTrue((bool)result);
        }
    }
}
