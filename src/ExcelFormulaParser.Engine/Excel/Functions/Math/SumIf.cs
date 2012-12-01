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
            var criteria = arguments.ElementAt(1).Value;
            var retVal = 0d;
            if (arguments.Count() > 2)
            {
                var sumRange = arguments.ElementAt(2).Value as IEnumerable<FunctionArgument>;
                retVal = CalculateWithSumRange(args, criteria, sumRange);
            }
            else
            {
                retVal = CalculateSingleRange(args, criteria);
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double CalculateWithSumRange(IEnumerable<FunctionArgument> range, object criteria, IEnumerable<FunctionArgument> sumRange)
        {
            var retVal = 0d;
            var flattenedRange = ArgsToDoubleEnumerable(range);
            var flattenedSumRange = ArgsToDoubleEnumerable(sumRange);
            for (var x = 0; x < flattenedRange.Count(); x++)
            {
                var candidate = flattenedSumRange.ElementAt(x);
                if (_evaluator.Evaluate(flattenedRange.ElementAt(x), criteria.ToString()))
                {
                    retVal += Convert.ToDouble(candidate);
                }
            }
            return retVal;
        }

        private double CalculateSingleRange(IEnumerable<FunctionArgument> args, object expression)
        {
            var retVal = 0d;
            if (args != null)
            {
                foreach (var arg in args)
                {
                    retVal += Calculate(arg, expression);
                }
            }
            return retVal;
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
