﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Tests.TestHelpers;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Tests.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class InformationFunctionsTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_excelDataProvider);
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
            var result = _parser.Parse("ISNUMBER(10/2)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void IsErrorShouldReturnTrueWhenDivBy0()
        {
            var result = _parser.Parse("ISERROR(10/0)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void IsTextShouldReturnTrueWhenReferencedCellContainsText()
        {
            _excelDataProvider.Stub(x => x.GetRangeValues("A1")).Return(new List<ExcelCell> { new ExcelCell("abc", string.Empty, 0, 0) });
            var result = _parser.Parse("ISTEXT(A1)");
            Assert.IsTrue((bool)result);
        }
    }
}
