using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.VBA.Operators
{
    public interface IOperator
    {
        Operators Operator { get; }

        object Apply(object left, object right);

        int Precedence { get; }
    }
}
