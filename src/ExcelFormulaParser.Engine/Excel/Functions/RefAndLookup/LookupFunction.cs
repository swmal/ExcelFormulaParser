using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public abstract class LookupFunction : ExcelFunction
    {
        private readonly ValueMatcher _valueMatcher;

        public LookupFunction()
            : this(new ValueMatcher())
        {

        }

        public LookupFunction(ValueMatcher valueMatcher)
        {
            _valueMatcher = valueMatcher;
        }

        public override bool IsLookupFuction
        {
            get
            {
                return true;
            }
        }

        protected int IsMatch(object o1, object o2)
        {
            return _valueMatcher.IsMatch(o1, o2);
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
