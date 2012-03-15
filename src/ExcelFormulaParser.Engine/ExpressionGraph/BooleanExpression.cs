using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class BooleanExpression : AtomicExpression
    {
        public BooleanExpression(string expression)
            : base(expression)
        {

        }
        public override CompileResult Compile()
        {
            return new CompileResult(bool.Parse(ExpressionString), DataType.Boolean);
        }
    }
}
