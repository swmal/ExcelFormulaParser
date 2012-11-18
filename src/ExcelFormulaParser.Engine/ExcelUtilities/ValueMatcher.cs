using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class ValueMatcher
    {
        public virtual int IsMatch(object o1, object o2)
        {
            if (o1 != null && o2 == null) return 1;
            if (o1 == null && o2 != null) return -1;
            if (o1 == null && o2 == null) return 0;
            if (o1.GetType() == typeof(string) && o2.GetType() == typeof(string))
            {
                return o1.ToString().CompareTo(o2);
            }
            else if( o1.GetType() == typeof(string))
            {
                // todo... compare
            }
            else if (o2.GetType() == typeof(string))
            {
                // todo.. compare
            }
            return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
        }
    }
}
