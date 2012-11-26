using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using System.Text.RegularExpressions;
using ExcelFormulaParser.Engine.Utilities;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class Rows : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var range = ArgToString(arguments, 0);
            if (Regex.IsMatch(range, RegexConstants.ExcelAddress))
            {
                var factory = new RangeAddressFactory(context.ExcelDataProvider);
                var address = factory.Create(range);
                return CreateResult(address.ToRow - address.FromRow + 1, DataType.Integer);
            }
            throw new ArgumentException("Invalid range supplied");
        }
    }
}
