using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.IntegrationTests.ExcelDataProviderTests
{
    [TestClass]
    public class ExcelDataProviderIntegrationTests
    {
        [TestMethod]
        public void ShouldCallProviderInSumFunctionAndCalculateResult()
        {
            var expectedAddres = "A1:A2";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddres))
                .Return(new object[] { 1, 2 });
            var parser = new FormulaParser(provider);
            var result = parser.Parse(string.Format("sum({0})", expectedAddres));
            Assert.AreEqual(3d, result);
        }
    }
}
