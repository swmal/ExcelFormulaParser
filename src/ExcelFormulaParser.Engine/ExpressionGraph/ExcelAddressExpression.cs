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
        private readonly IndexToAddressTranslator _indexToAddressTranslator = new IndexToAddressTranslator();

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
            //_parsingContext.Ranges.CheckCircularReference(_parsingContext.Scopes.Current, rangeAddress);
            //_parsingContext.Ranges.Add(_parsingContext.Scopes.Current, rangeAddress);
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
                return factory.Create(rangeValueList.First());
            }
        }

        private List<object> HandleRangeValues(IEnumerable<ExcelDataItem> result)
        {
            var rangeValueList = new List<object>();
            for (int x = 0; x < result.Count(); x++)
            {
                var dataItem = result.ElementAt(x);
                var rangeValue = dataItem.Value;
                if (IsFormula(rangeValue))
                {
                    var address = RangeAddress.Create(dataItem.ColIndex, dataItem.RowIndex);
                    rangeValueList.Add(_parsingContext.Parser.Parse(rangeValue.ToString(), address));
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
