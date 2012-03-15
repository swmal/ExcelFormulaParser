﻿using System;
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
        private bool _negateNextExpression;

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
                if (token.TokenType == TokenType.Operator && OperatorsDict.Instance.TryGetValue(token.Value, out op))
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
                else if (token.TokenType == TokenType.Negator)
                {
                    _negateNextExpression = true;
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
            if (parent != null && token.TokenType == TokenType.Comma)
            {
                parent.AddChild(new GroupExpression());
                return;
            }
            if (_negateNextExpression)
            {
                token.Negate();
                _negateNextExpression = false;
            }
            var expression = _expressionFactory.Create(token);
            if (parent == null)
            {
                _graph.Add(expression);
            }
            else if (parent.IsFunctionExpression)
            {
                if (parent.Children.Count() == 0)
                {
                    var group = parent.AddChild(new GroupExpression());
                    group.AddChild(expression);
                }
                else
                {
                    parent.Children.Last().AddChild(expression);
                }
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenType == TokenType.String)
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
