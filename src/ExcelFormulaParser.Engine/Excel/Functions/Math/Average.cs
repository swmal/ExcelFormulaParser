using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Average : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double nValues = 0d, result = 0d;
            foreach (var arg in arguments)
            {
                Calculate(arg.Value, ref result, ref nValues);
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
            else if (arg is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg)
                {
                    Calculate(item.Value, ref retVal, ref nValues);
                }
            }
        }


    }
}
