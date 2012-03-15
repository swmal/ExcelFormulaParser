using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class Minute : TimeBaseFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateAndInitSerialNumber(arguments);
            var minutes = (int)System.Math.Round(GetMinute(SerialNumber), 0);
            return new CompileResult(minutes, DataType.Integer);
        }
    }
}
