using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.VBA.Operators
{
    public class OperatorsDict : Dictionary<string, IOperator>
    {
        public OperatorsDict()
        {
            Add("+", Operator.Plus);
            Add("-", Operator.Minus);
            Add("*", Operator.Multiply);
            Add("/", Operator.Divide);
            Add("&", Operator.Concat);
            Add("mod", Operator.Modulus);
        }

        public static IDictionary<string, IOperator> _instance;

        public static IDictionary<string, IOperator> Instance
        {
            get 
            {
                if (_instance == null)
                {
                    _instance = new OperatorsDict();
                }
                return _instance;
            }
        }
    }
}
