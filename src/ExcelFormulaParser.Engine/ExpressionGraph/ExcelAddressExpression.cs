using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExcelAddressExpression : Expression
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _context;

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext context)
            : base(expression)
        {
            if (excelDataProvider == null)
            {
                throw new ArgumentNullException("excelDataProvider");
            }
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            _excelDataProvider = excelDataProvider;
            _context = context;
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
            var rangeValueList = new List<object>();
            FormulaParser parser = null;
            for (int x = 0; x < result.Count(); x++ )
            {
                var rangeValue = result.ElementAt(x);
                if (rangeValue != null && rangeValue.ToString().StartsWith("="))
                {
                    if (parser == null) parser = new FormulaParser(_excelDataProvider);
                    rangeValueList.Add(parser.Parse(rangeValue.ToString()));
                    continue;
                }
                rangeValueList.Add(rangeValue);
            }
            if (rangeValueList.Count > 1)
            {
                return new CompileResult(rangeValueList, DataType.Enumerable);
            }
            else
            {
                var factory = new CompileResultFactory();
                return factory.Create(result.First());
            }
        }
    }
}
