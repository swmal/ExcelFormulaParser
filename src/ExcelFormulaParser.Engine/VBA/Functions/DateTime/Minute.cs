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
            ValidateArguments(arguments, 1);
            var firstArg = arguments.ElementAt(0).ToString();
            if (arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
            {
                var result = TimeStringParser.Parse(firstArg);
                return CreateResult(GetMinuteFromSerialNumber(result), DataType.Integer);
            }
            ValidateAndInitSerialNumber(arguments);
            return CreateResult(GetMinuteFromSerialNumber(SerialNumber), DataType.Integer);
        }

        private int GetMinuteFromSerialNumber(double serialNumber)
        {
            return (int)System.Math.Round(GetMinute(serialNumber), 0);
        }
    }
}
