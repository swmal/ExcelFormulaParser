using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class TokenSeparatorProvider : ITokenSeparatorProvider
    {
        public TokenSeparatorProvider ()
	    {
            if(!_isInitialized)
            {
                Init();
            }
	    }
        private bool _isInitialized = false;

        private static void Init()
        {
            _tokens.Clear();
            _tokens.Add("+", new Token("+", TokenType.Operator));
            _tokens.Add("-", new Token("-", TokenType.Operator));
            _tokens.Add("*", new Token("*", TokenType.Operator));
            _tokens.Add("/", new Token("/", TokenType.Operator));
            _tokens.Add("&", new Token("&", TokenType.Operator));
            _tokens.Add(">", new Token(">", TokenType.Operator));
            _tokens.Add("<", new Token("<", TokenType.Operator));
            _tokens.Add("(", new Token("(", TokenType.OpeningBracket));
            _tokens.Add(")", new Token(")", TokenType.ClosingBracket));
            _tokens.Add("'", new Token("'", TokenType.String));
            _tokens.Add("\"", new Token("\"", TokenType.String));
            _tokens.Add(",", new Token(",", TokenType.Comma));
        }

        private static Dictionary<string, Token> _tokens = new Dictionary<string, Token>();

        IDictionary<string, Token> ITokenSeparatorProvider.Tokens
        {
            get { return _tokens; }
        }
    }
}
