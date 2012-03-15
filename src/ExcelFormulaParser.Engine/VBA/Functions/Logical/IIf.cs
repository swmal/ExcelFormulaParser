using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Logical
{
    public class IIf : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 3);
            var condition = ArgToBool(arguments, 0);
            var firstStatement = arguments.ElementAt(1);
            var secondStatement = arguments.ElementAt(2);
            var factory = new CompileResultFactory();
            return condition ? factory.Create(firstStatement) : factory.Create(secondStatement);
        }
    }
}
