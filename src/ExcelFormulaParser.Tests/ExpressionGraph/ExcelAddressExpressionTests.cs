using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExpressionGraph;
using Rhino.Mocks;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.ExpressionGraph
{
    [TestClass]
    public class ExcelAddressExpressionTests
    {
        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfExcelDataProviderIsNull()
        {
            var expression = new ExcelAddressExpression("A1", null);
        }

        [TestMethod]
        public void ShouldCallReturnResultFromProvider()
        {
            var expectedAddres = "A1";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddres))
                .Return(new object[] { 1 });

            var expression = new ExcelAddressExpression(expectedAddres, provider);
            var result = expression.Compile();
            Assert.AreEqual(1, ((IEnumerable<object>)result.Result).First());
        }
    }
}
