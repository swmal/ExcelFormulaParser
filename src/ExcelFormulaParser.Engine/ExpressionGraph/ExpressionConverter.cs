using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExpressionConverter : IExpressionConverter
    {
        public StringExpression ToStringExpression(Expression expression)
        {
            var result = expression.Compile();
            var newExp = new StringExpression(result.Result.ToString());
            newExp.Operator = expression.Operator;
            return newExp;
        }

        public Expression FromCompileResult(CompileResult compileResult)
        {
            switch (compileResult.DataType)
            {
                case DataType.Integer:
                    return new IntegerExpression(compileResult.Result.ToString());
                case DataType.String:
                    return new StringExpression(compileResult.Result.ToString());
                case DataType.Decimal:
                    return new DecimalExpression(compileResult.Result.ToString());
            }
            return null;
        }
    }
}
