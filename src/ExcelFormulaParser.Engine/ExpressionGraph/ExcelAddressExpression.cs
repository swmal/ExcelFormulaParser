using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExcelAddressExpression : Expression
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : base(expression)
        {
            if (excelDataProvider == null)
            {
                throw new ArgumentNullException("excelDataProvider");
            }
            if (parsingContext == null)
            {
                throw new ArgumentNullException("context");
            }
            _excelDataProvider = excelDataProvider;
            _parsingContext = parsingContext;
        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override CompileResult Compile()
        {
            var rangeAddress = RangeAddress.Parse(ExpressionString);
            _parsingContext.Ranges.CheckCircularReference(rangeAddress);
            _parsingContext.Ranges.Add(rangeAddress);
            var result = _excelDataProvider.GetRangeValues(ExpressionString);
            if (result == null || result.Count() == 0)
            {
                return null;
            }
            var rangeValueList = HandleRangeValues(result);
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

        private List<object> HandleRangeValues(IEnumerable<object> result)
        {
            var rangeValueList = new List<object>();
            for (int x = 0; x < result.Count(); x++)
            {
                var rangeValue = result.ElementAt(x);
                if (IsFormula(rangeValue))
                {
                    rangeValueList.Add(_parsingContext.Parser.Parse(rangeValue.ToString()));
                }
                else
                {
                    rangeValueList.Add(rangeValue);
                }
            }
            return rangeValueList;
        }

        private static bool IsFormula(object rangeValue)
        {
            return rangeValue != null && rangeValue.ToString().StartsWith("=");
        }
    }
}
