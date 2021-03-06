﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathObj = System.Math;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Stdev : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var values = ArgsToDoubleEnumerable(arguments);
            return CreateResult(StandardDeviation(values), DataType.Decimal);
        }

        private static double StandardDeviation(IEnumerable<double> values)
        {
            double ret = 0;
            if (values.Count() > 0)
            {
                //Compute the Average       
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2       
                double sum = values.Sum(d => MathObj.Pow(d - avg, 2));
                //Put it all together       
                ret = MathObj.Sqrt((sum) / (values.Count() - 1));
            }
            return ret;
        } 

    }
}
