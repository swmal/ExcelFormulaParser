using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Subtotal : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var calcType = ArgToInt(arguments, 0);
            var actualArgs = arguments.Skip(1);
            ExcelFunction function = null;
            switch(calcType)
            {
                case 1 :
                    function = new Average();
                    break;
                case 2:
                    function = new Count();
                    break;
                case 3:
                    function = new CountA();
                    break;
                case 4:
                    function = new Max();
                    break;
                case 5:
                    function = new Min();
                    break;
                case 6:
                    function = new Product();
                    break;
                case 7:
                    function = new Stdev();
                    break;
                case 8:
                    function = new StdevP();
                    break;
                case 9:
                    function = new Sum();
                    break;
                case 10:
                    function = new Var();
                    break;
                case 11:
                    function = new VarP();
                    break;
            }
            return function.Execute(actualArgs, context);
            //throw new NotImplementedException();
        }
    }
}
