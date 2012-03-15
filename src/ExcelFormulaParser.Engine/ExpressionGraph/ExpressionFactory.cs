﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExpressionFactory : IExpressionFactory
    {
        public Expression Create(Token token)
        {
            switch (token.TokenType)
            {
                case TokenType.Integer:
                    return new IntegerExpression(token.Value);
                case TokenType.String:
                    return new StringExpression(token.Value);
                case TokenType.Decimal:
                    return new DecimalExpression(token.Value);
                case TokenType.Boolean:
                    return new BooleanExpression(token.Value);
                case TokenType.ExcelAddress:
                    return new ExcelRangeExpression(token.Value);
                default:
                    return new StringExpression(token.Value);
            }
        }
    }
}
