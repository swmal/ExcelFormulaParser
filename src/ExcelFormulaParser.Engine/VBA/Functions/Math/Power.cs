using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Power : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var power = ArgToDecimal(arguments, 1);
            var result = System.Math.Pow((double)number, (double)power);
            return new CompileResult((decimal)result, DataType.Decimal);
        }
    }
}
