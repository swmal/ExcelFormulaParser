using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Exceptions
{
    public class CircularReferenceException : Exception
    {
        public CircularReferenceException(string message)
            : base(message)
        {

        }
    }
}
