using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class ValueMatcher
    {
        public const int IncompatibleOperands = -2;

        public virtual int IsMatch(object o1, object o2)
        {
            if (o1 != null && o2 == null) return 1;
            if (o1 == null && o2 != null) return -1;
            if (o1 == null && o2 == null) return 0;
            if (o1 is string && o2 is string)
            {
                return o1.ToString().CompareTo(o2);
            }
            else if( o1.GetType() == typeof(string))
            {
                return CompareStringToObject(o1.ToString(), o2);
            }
            else if (o2.GetType() == typeof(string))
            {
                return CompareObjectToString(o1, o2.ToString());
            }
            return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
        }

        private int CompareStringToObject(string o1, object o2)
        {
            double d1;
            if (double.TryParse(o1, out d1))
            {
                return d1.CompareTo(Convert.ToDouble(o2));
            }
            return IncompatibleOperands;
        }

        private int CompareObjectToString(object o1, string o2)
        {
            double d2;
            if (double.TryParse(o2, out d2))
            {
                return Convert.ToDouble(o1).CompareTo(d2);
            }
            return IncompatibleOperands;
        }
    }
}
