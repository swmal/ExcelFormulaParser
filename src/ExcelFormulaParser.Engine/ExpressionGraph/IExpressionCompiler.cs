using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public interface IExpressionCompiler
    {
        object Compile(IEnumerable<Expression> expressions);
    }
}
