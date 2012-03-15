using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.VBA.Operators
{
    public enum Operators
    {
        Undefined,
        Concat,
        Plus,
        Minus,
        Multiply,
        Divide,
        Modulus,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual
    }
}
