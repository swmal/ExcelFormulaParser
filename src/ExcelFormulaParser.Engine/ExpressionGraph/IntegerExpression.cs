using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class IntegerExpression : AtomicExpression
    {
        public IntegerExpression(string expression)
            : base(expression)
        {

        }

        public override CompileResult Compile()
        {
            return new CompileResult(double.Parse(ExpressionString), DataType.Integer);
        }
    }
}
