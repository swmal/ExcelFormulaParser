using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Utilities
{
    public static class ExtensionMethods
    {
        public static void IsNotNullOrEmpty(this ArgumentInfo<string> val)
        {
            if (string.IsNullOrEmpty(val.Value))
            {
                throw new ArgumentException(val.Name + " cannot be null or empty");
            }
        }

        public static void IsNotNull<T>(this ArgumentInfo<T> val)
            where T : class
        {
            if (val.Value == null)
            {
                throw new ArgumentNullException(val.Name);
            }
        }
    }
}
