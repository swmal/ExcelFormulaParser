using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public abstract class AtomicExpression : Expression
    {
        public AtomicExpression(string expression)
            : base(expression)
        {

        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }
    }
}
