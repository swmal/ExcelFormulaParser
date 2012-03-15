using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class RandBetween : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 2);
            var low = ArgToDecimal(arguments, 0);
            var high = ArgToDecimal(arguments, 1);
            var rand = new Rand().Execute(new object[0]).Result;
            var randPart = (CalulateDiff(high, low) * (double)rand) + 1;
            randPart = System.Math.Floor(randPart);
            return CreateResult(low + randPart, DataType.Integer);
        }

        private double CalulateDiff(double high, double low)
        {
            if (high > 0 && low < 0)
            {
                return high + low * - 1;
            }
            else if (high < 0 && low < 0)
            {
                return high * -1 - low * -1;
            }
            return high - low;
        }
    }
}
