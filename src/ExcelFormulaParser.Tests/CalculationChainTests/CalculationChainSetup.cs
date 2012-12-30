using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;
using ExcelFormulaParser.Engine.CalculationChain;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.Utilities;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Tests.CalculationChainTests
{
    [TestClass]
    public class CalculationChainSetup
    {
        private ExcelDataProvider _provider;
        private CalcChainContextBuilder _builder;
        private CalcChainContext _chainContext;
        private ParsingContext _parsingContext;
        private IdProvider _idProvider;

        [TestInitialize]
        public void Setup()
        {
            _idProvider = new IntegerIdProvider();
            _provider = MockRepository.GenerateStub<ExcelDataProvider>();
            SetupExcelProvider();
            _idProvider = new IntegerIdProvider();
            _chainContext = CalcChainContext.Create(_idProvider);
            _builder = new CalcChainContextBuilder();
            _parsingContext = ParsingContext.Create();
            _parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider);
            _parsingContext.Configuration.SetIdProvider(_idProvider);
            _parsingContext.Configuration.SetLexer(new Lexer(_parsingContext.Configuration.FunctionRepository, _parsingContext.NameValueProvider));
            _parsingContext.ExcelDataProvider = _provider;
        }

        private void SetupExcelProvider()
        {
            _provider.Stub(x => x.GetCellValue("A1")).Return(CreateFormulaCell("SUM(B1:B2) + A2"));
            _provider.Stub(x => x.GetCellValue("A2")).Return(CreateFormulaCell("SUM(C2:C3) + B2"));
            _provider.Stub(x => x.GetCellValue("B1")).Return(CreateFormulaCell("SUM(C1:C2)"));
            _provider.Stub(x => x.GetCellValue("B2")).Return(CreateFormulaCell("C4"));
            _provider.Stub(x => x.GetCellValue("C1")).Return(CreateValueCell(3));
            _provider.Stub(x => x.GetCellValue("C2")).Return(CreateValueCell(4));
            _provider.Stub(x => x.GetCellValue("C3")).Return(CreateValueCell(5));
            _provider.Stub(x => x.GetCellValue("C4")).Return(CreateFormulaCell("D1"));
            _provider.Stub(x => x.GetCellValue("D1")).Return(CreateValueCell(4));

            _provider.Stub(x => x.GetWorksheetFormulas("ws1")).Return(
                new Dictionary<string, string>
                {
                    {"A1", "SUM(B1:B2) + A2"},
                    {"A2", "SUM(C2:C3) + B2"},
                    {"B1", "SUM(C1:C2)"},
                    {"B2", "C4"},
                    {"C4", "D1"}
                }
            );
        }

        private ExcelCell CreateFormulaCell(string formula)
        {
            return CreateCell(formula, null, 0, 0);
        }

        private ExcelCell CreateValueCell(object value)
        {
            return CreateCell(null, value, 0, 0);
        }

        private ExcelCell CreateCell(string formula, object value, int row, int col)
        {
            return new ExcelCell(value, formula, col, row);
        }

        [TestMethod]
        public void ShouldResultIn4Chains()
        {
            var context = _builder.Build(_parsingContext, "ws1");
            Assert.AreEqual(4, context.CalcChains.Count());
        }

    }
}
