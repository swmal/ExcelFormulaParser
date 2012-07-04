using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Mod : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var n1 = ArgToDecimal(arguments, 0);
            var n2 = ArgToDecimal(arguments, 1);
            return new CompileResult(n1 % n2, DataType.Decimal);
        }
    }
}
