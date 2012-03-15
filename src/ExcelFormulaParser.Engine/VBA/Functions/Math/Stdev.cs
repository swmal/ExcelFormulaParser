using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathObj = System.Math;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Stdev : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var values = ArgsToDoubleEnumerable(arguments);
            return CreateResult(StandardDeviation(values), DataType.Decimal);
        }

        private static double StandardDeviation(IEnumerable<double> values)
        {
            double avg = values.Average();
            return MathObj.Sqrt(values.Average(v => MathObj.Pow(v - avg, 2)));
        }
    }
}
