using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Information
{
    public class IsNumber : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.ElementAt(0).Value;
            if (arg is double || arg is int)
            {
                return CreateResult(true, DataType.Boolean);
            }
            return CreateResult(false, DataType.Boolean);
        }
    }
}
