using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class TokenizerContext
    {
        public TokenizerContext(string formula)
        {
            if (!string.IsNullOrEmpty(formula))
            {
                _chars = formula.ToArray();
            }
            _result = new List<Token>();
            _currentToken = new StringBuilder();
        }

        private char[] _chars;
        private List<Token> _result;
        private StringBuilder _currentToken;

        public char[] FormulaChars
        {
            get { return _chars; }
        }

        public IEnumerable<Token> Result
        {
            get { return _result; }
        }

        public bool IsInString
        {
            get;
            private set;
        }

        public void ToggleIsInString()
        {
            IsInString = !IsInString;
        }

        public string CurrentToken
        {
            get { return _currentToken.ToString(); }
        }

        public void NewToken()
        {
            _currentToken = new StringBuilder();
        }

        public void AddToken(Token token)
        {
            _result.Add(token);
        }

        public void AppendToCurrentToken(char c)
        {
            _currentToken.Append(c.ToString());
        }
    }
}
