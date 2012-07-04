﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Product : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var result = 0d;
            var index = 0;
            while (result == 0d && index < arguments.Count())
            {
                result = CalculateFirstItem(arguments, index++);
            }
            result = CalculateCollection(arguments.Skip(index), result, (arg, current) =>
            {
                if (ShouldIgnore(arg)) return current;
                var obj = arg.Value;
                if (obj != null)
                {
                    if (obj is double)
                    {
                        current *= (double)obj;
                    }
                    else if (obj is int)
                    {
                        current *= (int)obj;
                    }
                }
                return current;
            });
            return CreateResult(result, DataType.Decimal);
        }

        private double CalculateFirstItem(IEnumerable<FunctionArgument> arguments, int index)
        {
            var element = arguments.ElementAt(index);
            var argList = new List<FunctionArgument> { element };
            var valueList = ArgsToDoubleEnumerable(argList);
            var result = 0d;
            foreach (var value in valueList)
            {
                if (result == 0d && value > 0d)
                {
                    result = value;
                }
                else
                {
                    result *= value;
                }
            }
            return result;
        }

        //private double CalculateFirstItem2(IEnumerable<FunctionArgument> arguments, int index)
        //{
        //    var firstElement = arguments.ElementAt(index);
        //    if (ShouldIgnore(firstElement))
        //    {
        //        return 0d;
        //    }
        //    var elementValue = firstElement.Value;
        //    if (elementValue is IEnumerable<FunctionArgument>)
        //    {
        //        var items = (IEnumerable<FunctionArgument>)elementValue;
        //        double? result = null;
        //        foreach (var item in items)
        //        {
        //            if (ShouldIgnore(item))
        //            {
        //                continue;
        //            }
        //            if (item.Value is double)
        //            {
        //                if (result.HasValue)
        //                {
        //                    result *= (double)item.Value;
        //                }
        //                else
        //                {
        //                    result = (double)item.Value;
        //                }
        //            }
        //        }
        //        return result ?? 0d;
        //    }
        //    return ArgToDecimal(arguments, 0);
        //}
    }
}
