using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class CountA : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var nItems = 0d;
            Calculate(arguments, ref nItems);
            return CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ref double nItems)
        {
            foreach (var item in items)
            {
                if (item.Value is IEnumerable<FunctionArgument>)
                {
                    Calculate((IEnumerable<FunctionArgument>)item.Value, ref nItems);
                }
                else if (ShouldCount(item.Value))
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
