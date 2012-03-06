using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public class CStr : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            if (arguments == null)
            {
                throw new ArgumentNullException("arguments");
            }
            return new CompileResult(arguments.First().ToString(), DataType.String);
        }
    }
}
