using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Max : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var values = ArgsToDoubleEnumerable(arguments);
            return CreateResult(values.Max(), DataType.Decimal);
        }
    }
}
