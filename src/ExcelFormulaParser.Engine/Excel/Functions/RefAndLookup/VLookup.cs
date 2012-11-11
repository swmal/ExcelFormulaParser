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
            var searchedValue = arguments.ElementAt(0).Value;
            var address = arguments.ElementAt(1).Value.ToString();
            var colIndex = ArgToInt(arguments, 2);
            var factory = new RangeAddressFactory(context.ExcelDataProvider);
            var rangeAddress = factory.Create(address);
            var fromRow = rangeAddress.FromRow;
            var col = rangeAddress.FromCol;
            for (var row = fromRow; row <= rangeAddress.ToRow; row++)
            {
                var cellValue = context.ExcelDataProvider.GetCellValue(row, col);
                if (cellValue != null && IsMatch(cellValue.Value, searchedValue))
                {
                    var correspondingVal = context.ExcelDataProvider.GetCellValue(row, col + colIndex);
                    if (correspondingVal != null)
                    {
                        return CreateResult(correspondingVal.Value, DataType.String);
                    }
                }
            }
            return CreateResult(null, DataType.String);
        }
    }
}
