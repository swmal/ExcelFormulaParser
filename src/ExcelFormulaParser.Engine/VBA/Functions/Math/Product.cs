using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Product : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 2);
            var first = CalculateFirstItem(arguments);
            var result = CalculateCollection(arguments.Skip(1), first, (obj, current) =>
            {
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

        private double CalculateFirstItem(IEnumerable<object> arguments)
        {
            var firstElement = arguments.ElementAt(0);
            if (firstElement is IEnumerable<object>)
            {
                var items = (IEnumerable<object>)firstElement;
                double? result = null;
                foreach (var item in items)
                {
                    if (item is double)
                    {
                        if (result.HasValue)
                        {
                            result *= (double)item;
                        }
                        else
                        {
                            result = (double)item;
                        }
                    }
                }
                return result ?? 0d;
            }
            return ArgToDecimal(arguments, 0);
        }
    }
}
