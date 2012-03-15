using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Text
{
    public class Left : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 2);
            var str = ArgToString(arguments, 0);
            var length = ArgToInt(arguments, 1);
            return CreateResult(str.Substring(0, length), DataType.String);
        }
    }
}
