using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class DecimalExpression : AtomicExpression
    {
        public DecimalExpression(string expression)
            : base(expression)
        {
            
        }

        public override CompileResult Compile()
        {
            string exp = string.Empty;
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            if (decimalSeparator == ",")
            {
                
                exp = ExpressionString.Replace('.', ',');
            }
            return new CompileResult(decimal.Parse(exp), DataType.Decimal);
        }
    }
}
