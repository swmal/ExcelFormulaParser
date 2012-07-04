using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Rand : ExcelFunction
    {
        private static int Seed
        {
            get;
            set;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            Seed = Seed > 50 ? 0 : Seed + 5;
            var val = new Random(System.DateTime.Now.Millisecond + Seed).NextDouble();
            return CreateResult(val, DataType.Decimal);
        }
    }
}
