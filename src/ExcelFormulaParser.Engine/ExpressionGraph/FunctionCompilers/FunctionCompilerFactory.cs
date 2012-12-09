using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Excel.Functions;

namespace ExcelFormulaParser.Engine.ExpressionGraph.FunctionCompilers
{
    public class FunctionCompilerFactory
    {
        public virtual FunctionCompiler Create(ExcelFunction function)
        {
            if (function.IsLookupFuction) return new LookupFunctionCompiler(function);
            if (function.IsErrorHandlingFunction) return new ErrorHandlingFunctionCompiler(function);
            return new DefaultCompiler(function);
        }
    }
}
