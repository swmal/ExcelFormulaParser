﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Operators
{
    public interface IOperator
    {
        Operators Operator { get; }

        CompileResult Apply(CompileResult left, CompileResult right);

        int Precedence { get; }
    }
}
