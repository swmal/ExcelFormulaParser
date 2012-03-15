using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Sqrt : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var arg = ArgToDecimal(arguments, 0);
            var result = System.Math.Sqrt((double)arg);
            return CreateResult((double)result, DataType.Decimal);
        }
    }
}
