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
            System.DateTime date = System.DateTime.MinValue;
            if (dateObj is double)
            {
                date = System.DateTime.FromOADate((double)dateObj);
            }
            if (dateObj is string)
            {
                date = System.DateTime.Parse(dateObj.ToString());
            }
            return CreateResult(date.Day, DataType.Integer);
        }
    }
}
