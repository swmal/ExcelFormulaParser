using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Exceptions
{
    public class ExcelFunctionException : Exception
    {
        public ExcelFunctionException(string message, ExcelErrorCodes errorCode)
            : base(message)
        {
            ErrorCode = errorCode.Code;
        }

        public ExcelFunctionException(string message, Exception innerException, ExcelErrorCodes errorCode)
            : base(message, innerException)
        {
            ErrorCode = errorCode.Code;
        }

        public string ErrorCode
        {
            get;
            private set;
        }
    }
}
