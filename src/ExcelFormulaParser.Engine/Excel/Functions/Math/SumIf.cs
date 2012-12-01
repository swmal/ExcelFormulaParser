using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class SumIf : HiddenValuesHandlingFunction
    {
        private readonly NumericExpressionEvaluator _evaluator;

        public SumIf()
            : this(new NumericExpressionEvaluator())
        {

        }

        public SumIf(NumericExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _evaluator = evaluator;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var args = arguments.ElementAt(0).Value as IEnumerable<FunctionArgument>;
            var expression = arguments.ElementAt(1).Value;
            var retVal = 0d;
            if (args != null)
            {
                foreach (var arg in args)
                {
                    retVal += Calculate(arg, expression);
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double Calculate(FunctionArgument arg, object expression)
        {
            var retVal = 0d;
            if (ShouldIgnore(arg) || !_evaluator.Evaluate(arg.Value, expression.ToString()))
            {
                return retVal;
            }
            if (arg.Value is double || arg.Value is int)
            {
                retVal += Convert.ToDouble(arg.Value);
            }
            else if (arg.Value is System.DateTime)
            {
                retVal += Convert.ToDateTime(arg.Value).ToOADate();
            }
            else if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    retVal += Calculate(item, expression);
                }
            }
            return retVal;
        }
    }
}
