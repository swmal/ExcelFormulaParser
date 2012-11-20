using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine.ExpressionGraph;

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

        protected CompileResult Lookup(LookupNavigator navigator, LookupArguments lookupArgs)
        {
            object lastValue = null;
            object lastLookupValue = null;
            do
            {
                var matchResult = IsMatch(navigator.CurrentValue, lookupArgs.SearchedValue);
                if (matchResult == 0)
                {
                    return CreateResult(navigator.GetLookupValue(), DataType.String);
                }
                if (lookupArgs.RangeLookup)
                {
                    if (lastValue != null && matchResult > 0)
                    {
                        return CreateResult(lastLookupValue, DataType.String);
                    }
                    lastValue = navigator.CurrentValue;
                    lastLookupValue = navigator.GetLookupValue();
                }
            }
            while (navigator.MoveNext());

            return CreateResult(null, DataType.String);
        }
    }
}
