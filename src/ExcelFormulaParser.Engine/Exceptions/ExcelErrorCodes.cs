using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Exceptions
{
    public class ExcelErrorCodes
    {
        private ExcelErrorCodes(string code)
        {
            Code = code;
        }

        public string Code
        {
            get;
            private set;
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }

        public override bool  Equals(object obj)
        {
            if (obj is ExcelErrorCodes)
            {
                return ((ExcelErrorCodes)obj).Code.Equals(Code);
            }
 	        return false;
        }

        public static bool operator == (ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return c1.Code.Equals(c2.Code);
        }

        public static bool operator !=(ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return !c1.Code.Equals(c2.Code);
        }

        public static ExcelErrorCodes Value
        {
            get { return new ExcelErrorCodes("#VALUE!"); }
        }

        public static ExcelErrorCodes Name
        {
            get { return new ExcelErrorCodes("#NAME?"); }
        }

        public static ExcelErrorCodes NoValueAvaliable
        {
            get { return new ExcelErrorCodes("#N/A"); }
        }
    }
}
