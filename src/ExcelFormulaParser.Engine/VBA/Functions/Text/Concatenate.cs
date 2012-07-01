using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Text
{
    public class Concatenate : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null)
            {
                return CreateResult(string.Empty, DataType.String);
            }
            var sb = new StringBuilder();
            foreach (var arg in arguments)
            {
                sb.Append(arg.Value.ToString());
            }
            return CreateResult(sb.ToString(), DataType.String);
        }
    }
}
