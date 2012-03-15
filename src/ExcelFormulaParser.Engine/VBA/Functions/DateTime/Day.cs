using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class Day : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var dateObj = arguments.ElementAt(0);
            var date = System.DateTime.FromOADate((double)dateObj);
            return new CompileResult(date.Day, DataType.Integer);
        }
    }
}
