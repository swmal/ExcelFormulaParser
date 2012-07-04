using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public class ArgumentCollectionUtil
    {
        public virtual IEnumerable<FunctionArgument> FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments)
        {
            var argList = new List<FunctionArgument>();
            FuncArgsToFlatEnumerable(arguments, argList);
            return argList;
        }

        public virtual void FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, List<FunctionArgument> argList)
        {
            foreach (var arg in arguments)
            {
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    FuncArgsToFlatEnumerable((IEnumerable<FunctionArgument>)arg.Value, argList);
                }
                else
                {
                    argList.Add(arg);
                }
            }
        }

        public virtual IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments)
        {
            var values = new List<double>();
            var args = FuncArgsToFlatEnumerable(arguments);
            for (var x = 0; x < args.Count(); x++)
            {
                var arg = args.ElementAt(x).Value;
                if (arg.GetType() == typeof(double) || arg.GetType() == typeof(int))
                {
                    values.Add(Convert.ToDouble(arg));
                }
            }
            return values;
        }

        public virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument, double, double> action)
        {
            foreach (var item in collection)
            {
                if (item.Value is IEnumerable<FunctionArgument>)
                {
                    result = CalculateCollection((IEnumerable<FunctionArgument>)item.Value, result, action);
                }
                else
                {
                    result = action(item, result);
                }
            }
            return result;
        }
    }
}
