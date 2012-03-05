using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public interface IExpressionFactory
    {
        Expression Create(Token token);
    }
}
