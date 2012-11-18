using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class VLookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var lookupArgs = new LookupArguments(arguments);

            var navigator = new LookupNavigator(LookupDirection.Vertical, lookupArgs, context.ExcelDataProvider);
            do
            {
                if (IsMatch(navigator.CurrentValue, lookupArgs.SearchedValue))
                {
                    return CreateResult(navigator.GetLookupValue(), DataType.String);
                }
            }
            while (navigator.MoveNext());

            return CreateResult(null, DataType.String);
        }
    }
}
