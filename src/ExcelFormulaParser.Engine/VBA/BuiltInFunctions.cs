using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Math;
using ExcelFormulaParser.Engine.VBA.Functions.Logical;

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
            _functions["lower"] = new Lower();
            _functions["upper"] = new Upper();
            _functions["left"] = new Left();
            _functions["right"] = new Right();
            // Numbers
            _functions["cint"] = new CInt();
            // Math
            _functions["power"] = new Power();
            _functions["sqrt"] = new Sqrt();
            // Logical
            _functions["if"] = new If();
            _functions["not"] = new Not();
        }

        private readonly Dictionary<string, VBAFunction> _functions;

        public IDictionary<string, VBAFunction> Functions
        {
            get { return _functions; }
        }
    }
}
