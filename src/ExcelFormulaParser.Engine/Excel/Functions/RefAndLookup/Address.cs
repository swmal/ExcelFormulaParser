using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class Address : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var row = ArgToInt(arguments, 0) - 1;
            var col = ArgToInt(arguments, 1) - 1;
            var referenceType = ExcelReferenceType.AbsoluteRowAndColumn;
            var worksheetSpec = string.Empty;
            if (arguments.Count() > 2)
            {
                referenceType = (ExcelReferenceType)ArgToInt(arguments, 2);
            }
            if (arguments.Count() > 3)
            {
                var fourthArg = arguments.ElementAt(3).Value;
                if(fourthArg.GetType().Equals(typeof(bool)) && !(bool)fourthArg)
                {
                    throw new InvalidOperationException("Excelformulaparser does not support the R1C1 format!");
                }
                if (fourthArg.GetType().Equals(typeof(string)))
                {
                    worksheetSpec = fourthArg.ToString() + "!";
                }
            }
            var translator = new IndexToAddressTranslator(context.ExcelDataProvider, referenceType);
            return CreateResult(worksheetSpec + translator.ToAddress(col, row), DataType.ExcelAddress);
        }
    }
}
