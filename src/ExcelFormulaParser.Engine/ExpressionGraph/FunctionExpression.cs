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

        public override Expression MergeWithNext()
        {
            Expression returnValue = null;
            if (Next != null && Operator != null)
            {
                var result = Operator.Apply(Compile(), Next.Compile());
                var expressionString = result.Result.ToString();
                var converter = new ExpressionConverter();
                returnValue = converter.FromCompileResult(result);
                if (Next != null)
                {
                    Operator = Next.Operator;
                }
                else
                {
                    Operator = null;
                }
                Next = Next.Next;
            }
            return returnValue;
        }
    }
}
