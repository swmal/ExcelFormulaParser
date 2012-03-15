using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using System.Globalization;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public abstract class VBAFunction
    {
        public abstract CompileResult Execute(IEnumerable<object> arguments);

        protected void ValidateArguments(IEnumerable<object> arguments, int minLength)
        {
            if (arguments == null)
            {
                throw new ArgumentNullException("arguments");
            }
            if (arguments.Count() < minLength)
            {
                throw new ArgumentException(string.Format("Expecting at least {0} arguments", minLength));
            }
        }

        protected int ArgToInt(IEnumerable<object> arguments, int index)
        {
            var obj = arguments.ElementAt(index);
            if(obj == null) throw new ArgumentNullException("expected argument (int) was null");
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
            if(!int.TryParse(obj.ToString(), out result))
            {
                throw new ArgumentException("Could not parse " + obj.ToString() + " to int");
            }
            return result;
        }

        protected string ArgToString(IEnumerable<object> arguments, int index)
        {
            var obj = arguments.ElementAt(index);
            return obj != null ? obj.ToString() : string.Empty;
        }

        protected double ArgToDecimal(IEnumerable<object> arguments, int index)
        {
            var obj = arguments.ElementAt(index);
            var str = obj != null ? obj.ToString() : string.Empty;
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            if (decimalSeparator == ",")
            {
                
                str = str.Replace('.', ',');
            }
            return double.Parse(str);
        }

        /// <summary>
        /// If the argument is a boolean value its value will be returned.
        /// If the argument is an integer value, true will be returned if its
        /// value is not 0, otherwise false.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        protected bool ArgToBool(IEnumerable<object> arguments, int index)
        {
            var obj = arguments.ElementAt(index) ?? string.Empty;
            bool result;
            if (bool.TryParse(obj.ToString(), out result))
            {
                return result;
            }
            int intResult;
            if (int.TryParse(obj.ToString(), out intResult))
            {
                return intResult != 0;
            }
            return result;
        }

        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message)
        {
            if (condition())
            {
                throw new ArgumentException(message);
            }
        }

        protected bool IsNumeric(object val)
        {
            if (val == null) return false;
            return val.GetType() == typeof(int) || val.GetType() == typeof(double) || val.GetType() == typeof(decimal);
        }

        protected IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<object> arguments)
        {
            var values = new List<double>();
            for (var x = 0; x < arguments.Count(); x++)
            {
                var arg = arguments.ElementAt(x);
                if(IsNumeric(arg))
                {
                    values.Add((double)ArgToDecimal(arguments, x));
                }
            }
            return values;
        }

        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message, params string[] formats)
        {
            message = string.Format(message, formats);
            ThrowArgumentExceptionIf(condition, message);
        }

        protected CompileResult CreateResult(object result, DataType dataType)
        {
            return new CompileResult(result, dataType);
        }
    }
}
