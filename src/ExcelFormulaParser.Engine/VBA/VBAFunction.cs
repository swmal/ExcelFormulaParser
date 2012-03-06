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

        protected void ValidateArguments(IEnumerable<object> arguments, int minLength)
        {
            if (arguments == null)
            {
                throw new ArgumentNullException("arguments");
            }
            if (arguments.Count() < minLength)
            {
                throw new ArgumentException(string.Format("Expecting at least {0} arguments", minLength));
            }
        }
    }
}
