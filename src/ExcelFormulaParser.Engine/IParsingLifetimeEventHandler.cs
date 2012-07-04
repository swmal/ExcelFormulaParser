using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine
{
    public interface IParsingLifetimeEventHandler
    {
        void ParsingCompleted();
    }
}
