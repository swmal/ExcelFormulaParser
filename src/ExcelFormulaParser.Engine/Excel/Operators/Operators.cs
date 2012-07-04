using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Operators
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
        Equals,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        IntegerDivision,
        Exponentiation
    }
}
