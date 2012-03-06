using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class FunctionExpression : AtomicExpression
    {
        public FunctionExpression(string expression)
            : base(expression)
        {

        }
        public override CompileResult Compile()
        {
            var args = new List<object>();
            var arg = Children.First().Compile().Result;
            args.Add(arg);
            var function = FunctionRepository.GetFunction(ExpressionString);
            return function.Execute(args);
        }
    }
}
