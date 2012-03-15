using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Math;
using ExcelFormulaParser.Engine.VBA.Functions.Logical;
using ExcelFormulaParser.Engine.VBA.Functions.DateTime;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public class BuiltInFunctions : IVBAModule
    {
        public BuiltInFunctions()
        {
            _functions = new Dictionary<string, VBAFunction>();
            // Text
            _functions["text"] = new CStr();
            _functions["len"] = new Len();
            _functions["lower"] = new Lower();
            _functions["upper"] = new Upper();
            _functions["left"] = new Left();
            _functions["right"] = new Right();
            _functions["mid"] = new Mid();
            _functions["replace"] = new Replace();
            _functions["substitute"] = new Substitute();
            // Numbers
            _functions["cint"] = new CInt();
            // Math
            _functions["power"] = new Power();
            _functions["sqrt"] = new Sqrt();
            _functions["pi"] = new Pi();
            _functions["ceiling"] = new Ceiling();
            // Logical
            _functions["if"] = new If();
            _functions["not"] = new Not();
            _functions["and"] = new And();
            _functions["or"] = new Or();
            // Date
            _functions["date"] = new Date();
            _functions["today"] = new Today();
            _functions["now"] = new Now();
            _functions["day"] = new Day();
            _functions["month"] = new Month();
            _functions["year"] = new Year();
            _functions["time"] = new Time();
            _functions["hour"] = new Hour();
            _functions["minute"] = new Minute();
            _functions["second"] = new Second();
        }

        private readonly Dictionary<string, VBAFunction> _functions;

        public IDictionary<string, VBAFunction> Functions
        {
            get { return _functions; }
        }
    }
}
