using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Utilities
{
    public static class RegexConstants
    {
        public const string SingleCellAddress = @"^(('[^/\\?*\[\]]{1,31}'|[A-Za-z_]{1,31})!)?[A-Z]{1,3}[1-9]{1}[0-9]{0,7}$";
        public const string ExcelAddress = @"^(('[^/\\?*\[\]]{1,31}'|[A-Za-z_]{1,31})!)?[A-Z]{1,3}[1-9]{1}[0-9]{0,7}(\:[A-Z]{1,3}[1-9]{1}[0-9]{0,7}){0,1}$";
        public const string Boolean = @"^(true|false)$";
        public const string Decimal = @"^[0-9]+\.[0-9]+$";
        public const string Integer = @"^[0-9]+$";
    }
}
