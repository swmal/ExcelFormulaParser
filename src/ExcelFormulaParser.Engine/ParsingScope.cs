﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine
{
    public class ParsingScope : IDisposable
    {
        private readonly ParsingScopes _parsingScopes;

        public ParsingScope(ParsingScopes parsingScope)
        {
            _parsingScopes = parsingScope;
        }

        public void Dispose()
        {
            _parsingScopes.KillScope(this);
        }
    }
}
