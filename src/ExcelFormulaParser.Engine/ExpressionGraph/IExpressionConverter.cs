using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public interface IExpressionConverter
    {
        StringExpression ToStringExpression(Expression expression);
        Expression FromCompileResult(CompileResult compileResult);
    }
}
