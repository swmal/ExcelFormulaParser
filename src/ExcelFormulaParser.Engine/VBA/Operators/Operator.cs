using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Operators
{
    public class Operator : IOperator
    {
        private const int PrecedenceMultiplyDevide = 3;
        private const int PrecedenceIntegerDivision = 4;
        private const int PrecedenceModulus = 5;
        private const int PrecedenceAddSubtract = 10;
        private const int PrecedenceConcat = 15;

        private Operator() { }

        private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, object> implementation)
        {
            _implementation = implementation;
            _precedence = precedence;
            _operator = @operator;
        }

        private readonly Func<CompileResult, CompileResult, object> _implementation;
        private int _precedence;
        private Operators _operator;

        int IOperator.Precedence
        {
            get { return _precedence; }
        }

        Operators IOperator.Operator
        {
            get { return _operator; }
        }

        public object Apply(CompileResult left, CompileResult right)
        {
            return _implementation(left, right);
        }

        public static IOperator Plus
        {
            get
            {
                return new Operator(Operators.Plus, PrecedenceAddSubtract, (l, r) =>
                {
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return ((int)l.Result) + ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Integer)
                    {
                        return ((decimal)l.Result) + ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Decimal)
                    {
                        return ((decimal)l.Result) + ((decimal)r.Result);
                    }
                    return 0;
                }); 
            }
        }

        public static IOperator Minus
        {
            get
            {
                return new Operator(Operators.Minus, PrecedenceAddSubtract, (l, r) =>
                {
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return ((int)l.Result) - ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Integer)
                    {
                        return ((decimal)l.Result) - ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Decimal)
                    {
                        return ((decimal)l.Result) - ((decimal)r.Result);
                    }
                    return 0;
                });
            }
        }

        public static IOperator Multiply
        {
            get
            {
                return new Operator(Operators.Multiply, PrecedenceMultiplyDevide, (l, r) =>
                {
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return ((int)l.Result) * ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Integer)
                    {
                        return ((decimal)l.Result) * ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Decimal)
                    {
                        return ((decimal)l.Result) * ((decimal)r.Result);
                    }
                    return 0;
                });
            }
        }

        public static IOperator Divide
        {
            get
            {
                return new Operator(Operators.Divide, PrecedenceMultiplyDevide, (l, r) =>
                {
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return ((int)l.Result) / ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Integer)
                    {
                        return ((decimal)l.Result) / ((int)r.Result);
                    }
                    if (l.DataType == DataType.Decimal && r.DataType == DataType.Decimal)
                    {
                        return ((decimal)l.Result) / ((decimal)r.Result);
                    }
                    return 0;
                });
            }
        }

        public static IOperator Concat
        {
            get
            {
                return new Operator(Operators.Concat, PrecedenceConcat, (l, r) =>
                    {
                        var lStr = l.Result != null ? l.Result.ToString() : string.Empty;
                        var rStr = r.Result != null ? r.Result.ToString() : string.Empty;
                        return string.Concat(lStr, rStr);
                    });
            }
        }

        public static IOperator Modulus
        {
            get
            {
                return new Operator(Operators.Modulus, PrecedenceModulus, (l, r) =>
                {
                    return (int)l.Result % (int)r.Result;
                });
            }
        }
    }
}
