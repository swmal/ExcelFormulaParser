using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class Hour : TimeBaseFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateAndInitSerialNumber(arguments);
            return new CompileResult((int)GetHour(SerialNumber), DataType.Integer);
        }
    }
}
