using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Text
{
    public class Lower : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            return CreateResult(arguments.First().ToString().ToLower(), DataType.String);
        }
    }
}
