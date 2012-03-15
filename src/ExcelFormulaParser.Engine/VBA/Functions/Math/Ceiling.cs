using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Ceiling : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var significance = ArgToDecimal(arguments, 1);
            ValidateNumberAndSign(number, significance);
            if (significance < 1 && significance > 0)
            {
                var floor = System.Math.Floor(number);
                var rest = number - floor;
                var nSign = (int)(rest / significance) + 1;
                return new CompileResult(floor + (nSign * significance), DataType.Decimal);
            }
            else if (significance == 1)
            {
                return new CompileResult(System.Math.Ceiling(number), DataType.Decimal);
            }
            else
            {
                var result = number - (number % significance) + significance;
                return new CompileResult(result, DataType.Decimal);
            }
        }

        private void ValidateNumberAndSign(decimal number, decimal sign)
        {
            if (number > 0m && sign < 0)
            {
                var values = string.Format("num: {0}, sign: {1}", number, sign);
                throw new InvalidOperationException("Ceiling cannot handle a negative significance when the number is positive" + values);
            }
        }
    }
}
