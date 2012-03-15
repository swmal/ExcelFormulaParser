﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Text
{
    public class Mid : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 3);
            var text = ArgToString(arguments, 0);
            var startIx = ArgToInt(arguments, 1);
            var length = ArgToInt(arguments, 2);
            var result = text.Substring(startIx, length);
            return new CompileResult(result, DataType.String);
        }
    }
}
