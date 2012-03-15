using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Average : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            double nValues = 0d, result = 0d;
            foreach (var arg in arguments)
            {
                Calculate(arg, ref result, ref nValues);
            }
            return CreateResult(result/nValues, DataType.Decimal);
        }

        private void Calculate(object arg, ref double retVal, ref double nValues)
        {
            if (arg is double)
            {
                nValues++;
                retVal += Convert.ToDouble(arg);
            }
            else if (arg is int)
            {
                nValues++;
                retVal += Convert.ToDouble((int)arg);
            }
            else if(arg is bool)
            {
                nValues++;
                retVal += (bool)arg ? 1 : 0;
            }
            else if (arg is IEnumerable<object>)
            {
                foreach (var item in (IEnumerable<object>)arg)
                {
                    Calculate(item, ref retVal, ref nValues);
                }
            }
        }


    }
}
