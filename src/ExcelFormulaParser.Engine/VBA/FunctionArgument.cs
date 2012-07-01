using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public class FunctionArgument
    {
        public FunctionArgument(object val)
        {
            Value = val;
        }

        public object Value { get; private set; }

        public Type Type
        {
            get { return Value != null ? Value.GetType() : null; }
        }
    }
}
