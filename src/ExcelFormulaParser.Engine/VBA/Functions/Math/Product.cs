using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Product : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var first = CalculateFirstItem(arguments);
            var result = CalculateCollection(arguments.Skip(1), first, (arg, current) =>
            {
                var obj = arg.Value;
                if (obj != null)
                {
                    if (obj is double)
                    {
                        current *= (double)obj;
                    }
                }
                return current;
            });
            return CreateResult(result, DataType.Decimal);
        }

        private double CalculateFirstItem(IEnumerable<FunctionArgument> arguments)
        {
            var firstElement = arguments.ElementAt(0).Value;
            if (firstElement is IEnumerable<FunctionArgument>)
            {
                var items = (IEnumerable<FunctionArgument>)firstElement;
                double? result = null;
                foreach (var item in items)
                {
                    if (item.Value is double)
                    {
                        if (result.HasValue)
                        {
                            result *= (double)item.Value;
                        }
                        else
                        {
                            result = (double)item.Value;
                        }
                    }
                }
                return result ?? 0d;
            }
            return ArgToDecimal(arguments, 0);
        }
    }
}
