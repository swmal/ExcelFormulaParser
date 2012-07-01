using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class FunctionExpression : AtomicExpression
    {
        public FunctionExpression(string expression, ParsingContext parsingContext)
            : base(expression)
        {
            _parsingContext = parsingContext;
        }

        private readonly ParsingContext _parsingContext;

        public override bool IsFunctionExpression
        {
            get
            {
                return true;
            }
        }

        public override CompileResult Compile()
        {
            var args = new List<object>();
            foreach (var child in Children)
            {
                var arg = child.Compile().Result;
                args.Add(arg);
            }
            var function = _parsingContext.Configuration.FunctionRepository.GetFunction(ExpressionString);
            return function.Execute(args, _parsingContext);
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
