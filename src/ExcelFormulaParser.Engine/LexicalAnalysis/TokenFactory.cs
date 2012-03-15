using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelFormulaParser.Engine.VBA;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class TokenFactory : ITokenFactory
    {
        public TokenFactory()
            : this(new TokenSeparatorProvider())
        {

        }

        public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
        }

        private readonly ITokenSeparatorProvider _tokenSeparatorProvider;

        public Token Create(IEnumerable<Token> tokens, string token)
        {
            Token tokenSeparator = null;
            if (_tokenSeparatorProvider.Tokens.TryGetValue(token, out tokenSeparator))
            {
                return tokenSeparator;
            }
            if (tokens.Any() && tokens.Last().TokenType == TokenType.String)
            {
                return new Token(token, TokenType.StringContent);
            }
            if (!string.IsNullOrEmpty(token))
            {
                token = token.Trim();
            }
            if (Regex.IsMatch(token, @"^[0-9]+\.[0-9]+$"))
            {
                return new Token(token, TokenType.Decimal);
            }
            if(Regex.IsMatch(token, @"^[0-9]+$"))
            {
                return new Token(token, TokenType.Integer);
            }
            if (Regex.IsMatch(token, @"^(true|false)$", RegexOptions.IgnoreCase))
            {
                return new Token(token, TokenType.Boolean);
            }
            if (Regex.IsMatch(token, @"^(('[^/\\?*\[\]]{1,31}'|[A-Za-z_]{1,31})!)?[A-Z]{1,2}[1-9]{1}[0-9]{0,6}(\:[A-Z]{1,2}[1-9]{1}[0-9]{0,6}){0,1}$"))
            {
                return new Token(token, TokenType.ExcelAddress);
            }
            if (FunctionRepository.Exists(token))
            {
                return new Token(token, TokenType.Function);
            }
            return new Token(token, TokenType.Unrecognized);

        }
    }
}
