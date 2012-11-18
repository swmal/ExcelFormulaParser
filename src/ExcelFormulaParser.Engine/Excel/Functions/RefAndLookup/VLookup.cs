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


            //var factory = new RangeAddressFactory(context.ExcelDataProvider);
            //var rangeAddress = factory.Create(lookupArgs.RangeAddress);
            //var fromRow = rangeAddress.FromRow;
            //var col = rangeAddress.FromCol;
            //for (var row = fromRow; row <= rangeAddress.ToRow; row++)
            //{
            //    var cellValue = context.ExcelDataProvider.GetCellValue(row, col);
            //    if (cellValue != null && IsMatch(cellValue.Value, lookupArgs.SearchedValue))
            //    {
            //        // colIndex is one-based
            //        var column = col + lookupArgs.ColumnIndex - 1;
            //        var correspondingVal = context.ExcelDataProvider.GetCellValue(row, column);
            //        if (correspondingVal != null)
            //        {
            //            return CreateResult(correspondingVal.Value, DataType.String);
            //        }
            //    }
            //}
            return CreateResult(null, DataType.String);
        }
    }
}
