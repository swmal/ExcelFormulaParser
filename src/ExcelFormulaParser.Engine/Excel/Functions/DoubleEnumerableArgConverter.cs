using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public class DoubleEnumerableArgConverter : CollectionFlattener<double>
    {
        public IEnumerable<double> ConvertArgs(IEnumerable<FunctionArgument> arguments)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
            {
                if (arg.Value is double || arg.Value is int)
                {
                    argList.Add(Convert.ToDouble(arg.Value));
                }
            });
        }
    }
}
