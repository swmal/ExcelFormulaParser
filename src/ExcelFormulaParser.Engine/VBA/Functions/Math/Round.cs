using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Round : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var nDigits = ArgToInt(arguments, 1);
            if (nDigits < 0)
            {
                nDigits *= -1;
                return CreateResult(number - (number % (System.Math.Pow(10, nDigits))), DataType.Integer); 
            }
            return CreateResult(System.Math.Round(number, nDigits), DataType.Decimal);
        }
    }
}
