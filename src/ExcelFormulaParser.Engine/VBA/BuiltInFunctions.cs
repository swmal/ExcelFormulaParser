using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Text;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public class BuiltInFunctions : IVBAModule
    {
        public BuiltInFunctions()
        {
            _functions = new Dictionary<string, VBAFunction>();
            // Text
            _functions["cstr"] = new CStr();
            _functions["len"] = new Len();
            // Numbers
            _functions["cint"] = new CInt();
        }

        private readonly Dictionary<string, VBAFunction> _functions;

        public IDictionary<string, VBAFunction> Functions
        {
            get { return _functions; }
        }
    }
}
