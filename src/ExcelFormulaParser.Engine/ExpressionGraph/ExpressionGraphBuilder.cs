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
                else if (token.TokenType == TokenType.OpeningBracket)
                {
                    _tokenIndex++;
                    BuildGroupExpression(tokens, parent);
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
            if (!IsWaste(token))
            {
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
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenType == TokenType.String)
            {
                return true;
            }
            return false;
        }

        private void BuildGroupExpression(IEnumerable<Token> tokens, Expression parent)
        {
            if (parent == null)
            {
                if (_graph.Current == null)
                {
                    _graph.Add(new GroupExpression());
                }
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
