using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public abstract class VBAFunction
    {
        public abstract CompileResult Execute(IEnumerable<object> arguments);
    }
}
