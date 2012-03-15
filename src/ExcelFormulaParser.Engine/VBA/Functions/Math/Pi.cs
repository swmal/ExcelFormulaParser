﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Pi : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            var result = System.Math.Round((decimal)System.Math.PI, 14);
            return new CompileResult(result, DataType.Decimal);
        }
    }
}
