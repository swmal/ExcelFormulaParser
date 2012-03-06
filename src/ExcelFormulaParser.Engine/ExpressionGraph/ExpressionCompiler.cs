﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Operators;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExpressionCompiler : IExpressionCompiler
    {
        private IEnumerable<Expression> _expressions;
        private IExpressionConverter _expressionConverter;

        public ExpressionCompiler()
            : this(new ExpressionConverter())
        {

        }

        public ExpressionCompiler(IExpressionConverter expressionConverter)
        {
            _expressionConverter = expressionConverter;
        }

        public CompileResult Compile(IEnumerable<Expression> expressions)
        {
            _expressions = expressions;
            return PerformCompilation();
        }

        private CompileResult PerformCompilation()
        {
            var compiledExpressions = HandleGroupedExpressions();
            while(compiledExpressions.Any(x => x.Operator != null))
            {
                var prec = FindLowestPrecedence();
                compiledExpressions = HandlePrecedenceLevel(prec);
            }
            return compiledExpressions.First().Compile();
        }

        private IEnumerable<Expression> HandleGroupedExpressions()
        {
            var first = _expressions.First();
            var groupedExpressions = _expressions.Where(x => x.IsGroupedExpression);
            foreach(var groupedExpression in groupedExpressions)
            {
                var result = groupedExpression.Compile();
                var newExp = _expressionConverter.FromCompileResult(result);
                newExp.Operator = groupedExpression.Operator;
                newExp.Prev = groupedExpression.Prev;
                newExp.Next = groupedExpression.Next;
                if (groupedExpression.Prev != null)
                {
                    groupedExpression.Prev.Next = newExp;
                }
                if (groupedExpression == first)
                {
                    first = newExp;
                }
            }
            return RefreshList(first);
        }

        private IEnumerable<Expression> HandlePrecedenceLevel(int precedence)
        {
            var first = _expressions.First();
            var expressionsToHandle = _expressions.Where(x => x.Operator != null && x.Operator.Precedence == precedence);
            foreach (var expression in expressionsToHandle)
            {
                if (expression.Operator.Operator == Operators.Concat)
                {
                    var newExp = _expressionConverter.ToStringExpression(expression);
                    newExp.Prev = expression.Prev;
                    newExp.Next = expression.Next;
                    if (expression.Prev != null)
                    {
                        expression.Prev.Next = newExp;
                    }
                    if (expression.Next != null)
                    {
                        expression.Next.Prev = newExp;
                    }
                    newExp.MergeWithNext();
                    if (expression == first)
                    {
                        first = newExp;
                    }
                }
                else
                {
                    var mergedExp = expression.MergeWithNext();
                    if (expression == first)
                    {
                        first = mergedExp;
                    }
                }
            }
            return RefreshList(first);
        }

        private int FindLowestPrecedence()
        {
            return _expressions.Where(x => x.Operator != null).Min(x => x.Operator.Precedence);
        }

        private IEnumerable<Expression> RefreshList(Expression first)
        {
            var resultList = new List<Expression>();
            var exp = first;
            resultList.Add(exp);
            while (exp.Next != null)
            {
                resultList.Add(exp.Next);
                exp = exp.Next;
            }
            _expressions = resultList;
            return resultList;
        }
    }
}
