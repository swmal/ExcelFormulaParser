using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Text;
using ExcelFormulaParser.Engine.VBA.Functions.Math;
using ExcelFormulaParser.Engine.VBA.Functions.Logical;
using ExcelFormulaParser.Engine.VBA.Functions.DateTime;
using ExcelFormulaParser.Engine.VBA.Functions.Numeric;

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
            _functions["concatenate"] = new Concatenate();
            // Numbers
            _functions["int"] = new CInt();
            // Math
            _functions["power"] = new Power();
            _functions["sqrt"] = new Sqrt();
            _functions["pi"] = new Pi();
            _functions["ceiling"] = new Ceiling();
            _functions["count"] = new Count();
            _functions["counta"] = new CountA();
            _functions["floor"] = new Floor();
            _functions["sum"] = new Sum();
            _functions["stdev"] = new Stdev();
            _functions["stdevp"] = new StdevP();
            _functions["exp"] = new Exp();
            _functions["max"] = new Max();
            _functions["min"] = new Min();
            _functions["average"] = new Average();
            _functions["round"] = new Round();
            _functions["rand"] = new Rand();
            _functions["randbetween"] = new RandBetween();
            _functions["var"] = new Var();
            _functions["varp"] = new VarP();
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
