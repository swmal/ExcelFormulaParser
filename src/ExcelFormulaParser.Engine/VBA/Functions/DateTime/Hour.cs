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
            ValidateArguments(arguments, 1);
            var firstArg = arguments.ElementAt(0).ToString();
            if (arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
            {
                var result = TimeStringParser.Parse(firstArg);
                return CreateResult(GetHourFromSerialNumber(result));
            }
            ValidateAndInitSerialNumber(arguments);
            return CreateResult(GetHourFromSerialNumber(SerialNumber));
        }

        private int GetHourFromSerialNumber(double serialNumber)
        {
            return (int)System.Math.Round(GetHour(serialNumber), 0);
        }

        private CompileResult CreateResult(int hour)
        {
            return new CompileResult(hour, DataType.Integer);
        }
    }
}
