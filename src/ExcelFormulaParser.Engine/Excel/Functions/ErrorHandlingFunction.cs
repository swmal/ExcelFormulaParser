using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public abstract class ErrorHandlingFunction : ExcelFunction
    {
        public override bool IsErrorHandlingFunction
        {
            get
            {
                return true;
            }
        }

        public abstract CompileResult HandleError(string errorCode);
    }
}
