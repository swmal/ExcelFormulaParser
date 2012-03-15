﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Math
{
    public class Rand : VBAFunction
    {
        private static int Seed
        {
            get;
            set;
        }

        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            Seed = Seed > 50 ? 0 : Seed + 5;
            var val = new Random(System.DateTime.Now.Millisecond + Seed).NextDouble();
            return CreateResult(val, DataType.Decimal);
        }
    }
}
