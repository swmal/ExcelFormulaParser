using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class StringExpression : AtomicExpression
    {
        public StringExpression(string expression)
            : base(expression)
        {

        }

        public override CompileResult Compile()
        {
            return new CompileResult(ExpressionString, DataType.String);
        }
    }
}
