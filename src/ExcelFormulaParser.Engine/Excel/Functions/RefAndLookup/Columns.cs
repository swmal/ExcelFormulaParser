using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;
using System.Text.RegularExpressions;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class Columns : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var range = ArgToString(arguments, 0);
            if (Regex.IsMatch(range, RegexConstants.ExcelAddress))
            {
                var factory = new RangeAddressFactory(context.ExcelDataProvider);
                var address = factory.Create(range);
                return CreateResult(address.ToCol - address.FromCol + 1, DataType.Integer);
            }
            throw new ArgumentException("Invalid range supplied");
        }
    }
}
