using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class ExpressionFactory : IExpressionFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;

        public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
        {
            _excelDataProvider = excelDataProvider;
            _parsingContext = context;
        }


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
                    return new ExcelAddressExpression(token.Value, _excelDataProvider, _parsingContext);
                default:
                    return new StringExpression(token.Value);
            }
        }
    }
}
