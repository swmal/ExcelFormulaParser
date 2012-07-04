using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Excel;
using ExcelFormulaParser.Engine.Excel.Functions;

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
            var args = new List<FunctionArgument>();
            var function = _parsingContext.Configuration.FunctionRepository.GetFunction(ExpressionString);
            function.BeforeInvoke(_parsingContext);
            foreach (var child in Children)
            {
                var arg = child.Compile().Result;
                BuildFunctionArguments(arg, args);
            }
            return function.Execute(args, _parsingContext);
        }

        private static void BuildFunctionArguments(object result, List<FunctionArgument> args)
        {
            if (result is IEnumerable<object>)
            {
                var argList = new List<FunctionArgument>();
                foreach (var arg in ((IEnumerable<object>)result))
                {
                    BuildFunctionArguments(arg, argList);
                }
                args.Add(new FunctionArgument(argList));
            }
            else
            {
                args.Add(new FunctionArgument(result));
            }
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
