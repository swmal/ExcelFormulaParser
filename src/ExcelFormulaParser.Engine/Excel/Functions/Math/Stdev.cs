using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathObj = System.Math;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Stdev : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
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
