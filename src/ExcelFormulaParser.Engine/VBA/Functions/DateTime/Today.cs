﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class Today : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            return CreateResult(System.DateTime.Today.ToOADate(), DataType.Date);
        }
    }
}