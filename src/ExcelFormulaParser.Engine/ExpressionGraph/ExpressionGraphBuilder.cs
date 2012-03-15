using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Operators;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExpressionGraphBuilder :IExpressionGraphBuilder
    {
        private readonly ExpressionGraph _graph = new ExpressionGraph();
        private readonly IExpressionFactory _expressionFactory;
        private int _tokenIndex = 0;

        public ExpressionGraphBuilder()
            : this(new ExpressionFactory())
        {

        }

        public ExpressionGraphBuilder(IExpressionFactory expressionFactory)
        {
            _expressionFactory = expressionFactory;
        }

        public ExpressionGraph Build(IEnumerable<Token> tokens)
        {
            _tokenIndex = 0;
            _graph.Reset();
            BuildUp(tokens, null);
            return _graph;
        }

        private void BuildUp(IEnumerable<Token> tokens, Expression parent)
        {
            while (_tokenIndex < tokens.Count())
            {
                var token = tokens.ElementAt(_tokenIndex);
                IOperator op = null;
                if (OperatorsDict.Instance.TryGetValue(token.Value, out op))
                {
                    SetOperatorOnExpression(parent, op);
                }
                else if (token.TokenType == TokenType.Function)
                {
                    _tokenIndex++;
                    BuildFunctionExpression(tokens, parent, token.Value);
                }
                else if (token.TokenType == TokenType.OpeningBracket)
                {
                    _tokenIndex++;
                    BuildGroupExpression(tokens, parent);
                    if (parent is FunctionExpression)
                    {
                        return;
                    }
                }
                else if (token.TokenType == TokenType.ClosingBracket)
                {
                    break;
                }
                else
                {
                    CreateAndAppendExpression(parent, token);
                }
                _tokenIndex++;
            }
        }

        private void CreateAndAppendExpression(Expression parent, Token token)
        {
            if (IsWaste(token)) return;
            var expression = _expressionFactory.Create(token);
            if (parent == null)
            {
                _graph.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenType == TokenType.String  || token.TokenType == TokenType.Comma)
            {
                return true;
            }
            return false;
        }

        private void BuildFunctionExpression(IEnumerable<Token> tokens, Expression parent, string funcName)
        {
            if (parent == null)
            {
                _graph.Add(new FunctionExpression(funcName));
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                var func = new FunctionExpression(funcName);
                parent.AddChild(func);
                BuildUp(tokens, func);
            }
        }

        private void BuildGroupExpression(IEnumerable<Token> tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new GroupExpression());
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                BuildUp(tokens, parent);
            }
        }

        private void SetOperatorOnExpression(Expression parent, IOperator op)
        {
            if (parent == null)
            {
                _graph.Current.Operator = op;
            }
            else
            {
                parent.Children.Last().Operator = op;
            }
        }
    }
}
