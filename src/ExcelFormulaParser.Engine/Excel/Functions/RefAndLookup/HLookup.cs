using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class HLookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var lookupArgs = new LookupArguments(arguments);
            var navigator = new LookupNavigator(LookupDirection.Horizontal, lookupArgs, context.ExcelDataProvider);
            return Lookup(navigator, lookupArgs);
        }
    }
}
