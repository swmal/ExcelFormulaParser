using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Var : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var args = ArgsToDoubleEnumerable(arguments);
            double avg = args.Average();
            double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
            var result = d / (args.Count() - 1);
            return new CompileResult(result, DataType.Decimal);
        }
    }
}
