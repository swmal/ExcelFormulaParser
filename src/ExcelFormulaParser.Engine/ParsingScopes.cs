using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine
{
    public class ParsingScopes
    {
        private readonly IParsingLifetimeEventHandler _lifetimeEventHandler;

        public ParsingScopes(IParsingLifetimeEventHandler lifetimeEventHandler)
        {
            _lifetimeEventHandler = lifetimeEventHandler;
        }
        private Stack<ParsingScope> _scopes = new Stack<ParsingScope>();

        public virtual ParsingScope NewScope()
        {
            var scope = new ParsingScope(this);
            _scopes.Push(scope);
            return scope;
        }

        public virtual ParsingScope Current
        {
            get { return _scopes.Count() > 0 ? _scopes.Peek() : null; }
        }

        public virtual void KillScope(ParsingScope parsingScope)
        {
            _scopes.Pop();
            if (_scopes.Count() == 0)
            {
                _lifetimeEventHandler.ParsingCompleted();
            }
        }
    }
}
