﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public interface IExpressionCompiler
    {
        CompileResult Compile(IEnumerable<Expression> expressions);
    }
}
