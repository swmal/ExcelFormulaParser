using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Operators
{
    public interface IOperator
    {
        Operators Operator { get; }

        object Apply(CompileResult left, CompileResult right);

        int Precedence { get; }
    }
}
