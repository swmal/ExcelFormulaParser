using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph.CompileStrategy
{
    public class DefaultCompileStrategy : CompileStrategy
    {
        public DefaultCompileStrategy(Expression expression)
            : base(expression)
        {

        }
        public override Expression Compile()
        {
            return _expression.MergeWithNext();
        }
    }
}
