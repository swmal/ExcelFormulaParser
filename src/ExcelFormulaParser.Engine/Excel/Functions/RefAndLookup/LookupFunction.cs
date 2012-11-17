using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public abstract class LookupFunction : ExcelFunction
    {
        public override bool IsLookupFuction
        {
            get
            {
                return true;
            }
        }

        protected bool IsMatch(object o1, object o2)
        {
            if (o1 == null && o2 != null) return false;
            if (o1 != null && o2 == null) return false;
            if (o1 == null && o2 == null) return true;
            return o1.ToString().Equals(o2.ToString());
        }

        protected double? ToNumeric(object o)
        {
            if (o == null) return default(double?);
            if (o is double) return (double)o;
            var str = o.ToString();
            double res;
            if (double.TryParse(str, out res))
            {
                return res;
            }
            return default(double?);
        }
    }
}
