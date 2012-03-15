using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class Second : TimeBaseFunction
    {
        public Second()
            : base()
        {

        }
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.ElementAt(0).ToString();
            if (arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
            {
                var result = TimeStringParser.Parse(firstArg);
                return CreateResult(GetSecondFromSerialNumber(result));
            }
            ValidateAndInitSerialNumber(arguments);
            return CreateResult(GetSecondFromSerialNumber(SerialNumber));
        }

        private int GetSecondFromSerialNumber(double serialNumber)
        {
            return (int)System.Math.Round(GetSecond(serialNumber), 0);
        }

        private CompileResult CreateResult(int second)
        {
            return new CompileResult(second, DataType.Integer);
        }
    }
}
