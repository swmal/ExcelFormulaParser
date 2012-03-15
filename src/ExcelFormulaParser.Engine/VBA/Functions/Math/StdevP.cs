using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class StdevP : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            var args = ArgsToDoubleEnumerable(arguments);
            return CreateResult(StdevPImpl(args), DataType.Decimal);
        }

        private static object StdevPImpl(IEnumerable<double> args)
        {
            var result = double.NaN;
            if (args != null)
            {
                double sum=0;
                double sum2=0;
                int count=0;
                foreach (var r in args)
                {
                    if (r != double.NaN)
                    {
                        sum += r;
                        sum2 += (r*r);
                        count++;
                    }
                }
                if (count > 0)
                {
                    result = System.Math.Sqrt(((count * sum2 - sum * sum) / (count * count)));
                }
                else
                {
                    result = double.NaN;
                }
            }
            return result;

        }
    }
}
