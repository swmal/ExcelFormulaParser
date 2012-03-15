using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Logical
{
    public class Not : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.ElementAt(0);
            bool result = false;
            if (arg is bool)
            {
                result = !((bool)arg);
            }
            else if (arg is int)
            {
                result = ((int)arg) == 0;
            }
            return new CompileResult(result, DataType.Boolean);
        }
    }
}
