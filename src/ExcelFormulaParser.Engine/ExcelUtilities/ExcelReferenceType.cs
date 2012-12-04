using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public enum ExcelReferenceType
    {
        AbsoluteRowAndColumn = 1,
        AbsoluteRowRelativeColumn = 2,
        RelativeRowAbsolutColumn = 3,
        RelativeRowAndColumn = 4
    }
}
