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
            return new CompileResult(result, DataType.Enumerable);
        }
    }
}
