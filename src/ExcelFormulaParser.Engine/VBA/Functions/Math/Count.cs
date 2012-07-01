using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Count : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments, ParsingContext context)
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
            if (item.GetType() == typeof(int)
                ||
                item.GetType() == typeof(double)
                ||
                item.GetType() == typeof(decimal)
                ||
                item.GetType() == typeof(System.DateTime))
            {
                return true;
            }
            return false;
        }
    }
}
