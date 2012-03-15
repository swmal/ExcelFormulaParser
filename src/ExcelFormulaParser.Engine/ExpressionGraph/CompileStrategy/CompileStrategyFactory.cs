using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Operators;

namespace ExcelFormulaParser.Engine.ExpressionGraph.CompileStrategy
{
    public class CompileStrategyFactory : ICompileStrategyFactory
    {
        public CompileStrategy Create(Expression expression)
        {
            if (expression.Operator.Operator == Operators.Concat)
            {
                return new StringConcatStrategy(expression);
            }
            else
            {
                return new DefaultCompileStrategy(expression);
            }
        }
    }
}
