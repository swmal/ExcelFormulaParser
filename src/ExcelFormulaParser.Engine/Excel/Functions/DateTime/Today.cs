using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.DateTime
{
    public class Today : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return CreateResult(System.DateTime.Today.ToOADate(), DataType.Date);
        }
    }
}
