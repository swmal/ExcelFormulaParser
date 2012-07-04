using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class CompileResult
    {
        public CompileResult(object result, DataType dataType)
        {
            Result = result;
            DataType = dataType;
        }
        public object Result
        {
            get;
            private set;
        }

        public DataType DataType
        {
            get;
            private set;
        }

        public bool IsNumeric
        {
            get { return DataType == Engine.ExpressionGraph.DataType.Decimal || DataType == Engine.ExpressionGraph.DataType.Integer; }
        }
    }
}
