using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Excel;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Engine.ExpressionGraph.FunctionCompilers;

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
        private readonly FunctionCompilerFactory _functionCompilerFactory = new FunctionCompilerFactory();

        public override bool IsFunctionExpression
        {
            get
            {
                return true;
            }
        }

        public override CompileResult Compile()
        {
            var function = _parsingContext.Configuration.FunctionRepository.GetFunction(ExpressionString);
            var compiler = _functionCompilerFactory.Create(function);
            return compiler.Compile(Children, _parsingContext);
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
