using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.CalculationChain;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;

namespace ExcelFormulaParser.Tests.CalculationChainTests
{
    [TestClass]
    public class CalculationChainBuilderTests
    {
        private ParsingContext _parsingContext;
        private ExcelDataProvider _excelDataProvider;
        private CalculationChainBuilder _builder;

        [TestInitialize]
        public void Setup()
        {
             _parsingContext = ParsingContext.Create();
             _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
             _parsingContext.ExcelDataProvider = _excelDataProvider;
            _builder = new CalculationChainBuilder(_parsingContext);
        }

        [TestMethod]
        public void ShouldUseFirstWorksheetNameWhenBuildIsCalledWithoutParams()
        {
            _excelDataProvider.Stub(x => x.GetWorksheetNames()).Return(new List<string> { "a", "b" });
            _excelDataProvider.Stub(x => x.GetWorksheetFormulas("a")).Return(new Dictionary<string, string>());
            var chain = _builder.Build();
            _excelDataProvider.AssertWasCalled(x => x.GetWorksheetFormulas("a"));
        }

        [TestMethod]
        public void ShouldNotCallGetWorksheetFormulasIfNoWorksheetExists()
        {
            _excelDataProvider.Stub(x => x.GetWorksheetNames()).Return(new List<string>());
            var chain = _builder.Build();
            Assert.IsNotNull(chain);
        }
    }
}
