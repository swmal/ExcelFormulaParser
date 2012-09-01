using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class TokenFactory : ITokenFactory
    {
        public TokenFactory(FunctionRepository functionRepository, NameValueProvider nameValueProvider)
            : this(new TokenSeparatorProvider(), nameValueProvider, functionRepository)
        {

        }

        public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider, NameValueProvider nameValueProvider, FunctionRepository functionRepository)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
            _functionRepository = functionRepository;
            _nameValueProvider = nameValueProvider;
        }

        private readonly ITokenSeparatorProvider _tokenSeparatorProvider;
        private readonly FunctionRepository _functionRepository;
        private readonly NameValueProvider _nameValueProvider;

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
            if (Regex.IsMatch(token, RegexConstants.Decimal))
            {
                return new Token(token, TokenType.Decimal);
            }
            if(Regex.IsMatch(token, RegexConstants.Integer))
            {
                return new Token(token, TokenType.Integer);
            }
            if (Regex.IsMatch(token, RegexConstants.Boolean, RegexOptions.IgnoreCase))
            {
                return new Token(token, TokenType.Boolean);
            }
            if (Regex.IsMatch(token, RegexConstants.ExcelAddress))
            {
                return new Token(token, TokenType.ExcelAddress);
            }
            if (_functionRepository.Exists(token))
            {
                return new Token(token, TokenType.Function);
            }
            if (_nameValueProvider.IsNamedValue(token))
            {
                return new Token(token, TokenType.NameValue);
            }
            return new Token(token, TokenType.Unrecognized);

        }
    }
}
