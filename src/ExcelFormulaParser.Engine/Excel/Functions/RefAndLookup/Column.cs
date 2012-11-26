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
    public class Column : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null || arguments.Count() == 0)
            {
                return CreateResult(context.Scopes.Current.Address.FromCol + 1, DataType.Integer);
            }
            var rangeAddress = ArgToString(arguments, 0);
            if (Regex.IsMatch(rangeAddress, RegexConstants.ExcelAddress))
            {
                var factory = new RangeAddressFactory(context.ExcelDataProvider);
                var address = factory.Create(rangeAddress);
                return CreateResult(address.FromCol + 1, DataType.Integer);
            }
            throw new ArgumentException("An invalid argument was supplied");
        }
    }
}
