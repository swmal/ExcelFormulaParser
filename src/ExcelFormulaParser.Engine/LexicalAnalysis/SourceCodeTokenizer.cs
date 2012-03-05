using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class SourceCodeTokenizer : ISourceCodeTokenizer
    {
        public SourceCodeTokenizer()
            : this(new TokenFactory(), new TokenSeparatorProvider())
        {

        }
        public SourceCodeTokenizer(ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
        {
            _tokenFactory = tokenFactory;
            _tokenProvider = tokenProvider;
        }

        private readonly ITokenSeparatorProvider _tokenProvider;
        private readonly ITokenFactory _tokenFactory;

        public IEnumerable<Token> Tokenize(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return Enumerable.Empty<Token>();
            }
            var context = new TokenizerContext(input);
            foreach (var c in context.FormulaChars)
            {
                Token tokenSeparator;
                if(CharIsTokenSeparator(c, out tokenSeparator))
                {
                    if (context.IsInString && tokenSeparator.TokenType != TokenType.String)
                    {
                        context.AppendToCurrentToken(c);
                        continue;
                    }
                    if (tokenSeparator.TokenType == TokenType.String)
                    {
                        context.ToggleIsInString();
                    }
                    if (context.CurrentTokenHasValue)
                    {
                        context.AddToken(CreateToken(context));
                    }
                    context.AddToken(tokenSeparator);
                    context.NewToken();
                    continue;
                }
                context.AppendToCurrentToken(c);
            }
            if (context.CurrentTokenHasValue)
            {
                context.AddToken(CreateToken(context));
            }
            return context.Result;
        }

        private Token CreateToken(TokenizerContext context)
        {
            return _tokenFactory.Create(context.Result, context.CurrentToken);
        }

        private bool CharIsTokenSeparator(char c, out Token token)
        {
            var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
            token = result ? token = _tokenProvider.Tokens[c.ToString()] : null;
            return result;
        }
    }
}
