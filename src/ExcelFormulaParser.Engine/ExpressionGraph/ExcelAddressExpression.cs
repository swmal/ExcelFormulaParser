using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExcelAddressExpression : Expression
    {
        private readonly ExcelDataProvider _excelDataProvider;

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider)
            : base(expression)
        {
            if (excelDataProvider == null)
            {
                throw new ArgumentNullException("excelDataProvider");
            }
            _excelDataProvider = excelDataProvider;
        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override CompileResult Compile()
        {

            var result = _excelDataProvider.GetRangeValues(ExpressionString);
            if (result == null || result.Count() == 0)
            {
                return null;
            }
            if (result.Count() > 1)
            {
                return new CompileResult(result, DataType.Enumerable);
            }
            else
            {
                var factory = new CompileResultFactory();
                return factory.Create(result.First());
            }
        }
    }
}
