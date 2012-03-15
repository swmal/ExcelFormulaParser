using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class CountA : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var nItems = 0d;
            Calculate(arguments, ref nItems);
            return CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<object> items, ref double nItems)
        {
            foreach (var item in items)
            {
                if (item is IEnumerable<object>)
                {
                    Calculate((IEnumerable<object>)item, ref nItems);
                }
                else if (ShouldCount(item))
                {
                    nItems++;
                }
                
            }
        }

        private bool ShouldCount(object item)
        {
            if (item == null) return false;
            return (!string.IsNullOrEmpty(item.ToString()));
        }
    }
}
