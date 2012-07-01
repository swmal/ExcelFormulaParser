using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Logical
{
    public class If : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var condition = ArgToBool(arguments, 0);
            var firstStatement = arguments.ElementAt(1).Value;
            var secondStatement = arguments.ElementAt(2).Value;
            var factory = new CompileResultFactory();
            return condition ? factory.Create(firstStatement) : factory.Create(secondStatement);
        }
    }
}
