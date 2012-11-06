using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class VLookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var searchedValue = arguments.ElementAt(0).Value;
            var lookupMatrix = arguments.ElementAt(1).Value as List<FunctionArgument>;
            var colIndex = ArgToInt(arguments, 2);
            for (var ix = 0; ix < lookupMatrix.Count; ix++)
            {
                //if (lookupMatrix[ix].Value[0].ToString() == searchedValue.ToString())
                //{
                //    return CreateResult(lookupMatrix[ix][colIndex], DataType.String);
                //}
            }
            return CreateResult(null, DataType.String);
        }
    }
}
