using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.Utilities;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class Lookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var searchedValue = arguments.ElementAt(0).Value;
            Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
            var firstAddress = arguments.ElementAt(1).Value.ToString();
            var factory = new RangeAddressFactory(context.ExcelDataProvider);
            var address = factory.Create(firstAddress);
            var nRows = address.ToRow - address.FromRow;
            var nCols = address.ToCol - address.FromCol;
            var lookupArgs = new LookupArguments(searchedValue, firstAddress, nCols + 1, true);
            var navigator = new LookupNavigator(LookupDirection.Vertical, lookupArgs, context.ExcelDataProvider);
            return Lookup(navigator, lookupArgs);
        }
    }
}
