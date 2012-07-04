using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public class IntArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            Require.That(obj).Named("argument").IsNotNull();
            int result;
            var objType = obj.GetType();
            if (objType == typeof(int))
            {
                return (int)obj;
            }
            if (objType == typeof(double) || objType == typeof(decimal))
            {
                return Convert.ToInt32(obj);
            }
            if (!int.TryParse(obj.ToString(), out result))
            {
                throw new ArgumentException("Could not parse " + obj.ToString() + " to int");
            }
            return result;
        }
    }
}
