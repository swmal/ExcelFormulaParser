using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.Utilities;
using ExcelFormulaParser.Engine.ExcelUtilities;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class Lookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            if (HaveTwoRanges(arguments))
            {
                return HandleTwoRanges(arguments, context);
            }
            return HandleSingleRange(arguments, context);
        }

        private bool HaveTwoRanges(IEnumerable<FunctionArgument> arguments)
        {
            if (arguments.Count() == 2) return false;
            return (Regex.IsMatch(arguments.ElementAt(2).Value.ToString(), RegexConstants.ExcelAddress));
        }

        private CompileResult HandleSingleRange(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments.ElementAt(0).Value;
            Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
            var firstAddress = arguments.ElementAt(1).Value.ToString();
            var address = GetRangeAddress(context.ExcelDataProvider, firstAddress);
            var nRows = address.ToRow - address.FromRow;
            var nCols = address.ToCol - address.FromCol;
            var lookupIndex = nCols + 1;
            var lookupDirection = LookupDirection.Vertical;
            if (nCols > nRows)
            {
                lookupIndex = nRows + 1;
                lookupDirection = LookupDirection.Horizontal;
            }
            var lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, 0, true);
            var navigator = new LookupNavigator(lookupDirection, lookupArgs, context.ExcelDataProvider);
            return Lookup(navigator, lookupArgs);
        }

        private CompileResult HandleTwoRanges(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments.ElementAt(0).Value;
            Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
            Require.That(arguments.ElementAt(2).Value).Named("secondAddress").IsNotNull();
            var firstAddress = arguments.ElementAt(1).Value.ToString();
            var secondAddress = arguments.ElementAt(2).Value.ToString();
            var address1 = GetRangeAddress(context.ExcelDataProvider, firstAddress);
            var address2 = GetRangeAddress(context.ExcelDataProvider, secondAddress);
            var nRows = address1.ToRow - address1.FromRow;
            var nCols = address1.ToCol - address1.FromCol;
            var lookupIndex = (address2.FromCol - address1.FromCol) + 1;
            var lookupOffset = address2.FromRow - address1.FromRow;
            var lookupDirection = LookupDirection.Vertical;
            if (nCols > nRows)
            {
                lookupIndex = (address2.FromRow - address1.FromRow) + 1;
                lookupOffset = address2.FromCol - address1.FromCol;
                lookupDirection = LookupDirection.Horizontal;
            }
            var lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, lookupOffset,  true);
            var navigator = new LookupNavigator(lookupDirection, lookupArgs, context.ExcelDataProvider);
            return Lookup(navigator, lookupArgs);
        }

        private static RangeAddress GetRangeAddress(ExcelDataProvider excelDataProvider, string address)
        {
            var factory = new RangeAddressFactory(excelDataProvider);
            return factory.Create(address);
        }
    }
}
