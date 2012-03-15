using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Sum : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            var retVal = 0d;
            if (arguments != null)
            {
                foreach (var arg in arguments)
                {
                    retVal += Calculate(arg);
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double Calculate(object arg)
        {
            var retVal = 0d;
            if (arg is double)
            {
                retVal += Convert.ToDouble((double)arg);
            }
            else if (arg is int)
            {
                retVal += Convert.ToDouble((int)arg);
            }
            else if (arg is System.DateTime)
            {
                retVal += Convert.ToDateTime(arg).ToOADate();
            }
            else if (arg is IEnumerable<object>)
            {
                foreach (var item in (IEnumerable<object>)arg)
                {
                    retVal += Calculate(item);
                }
            }
            return retVal;
        }
    }
}
