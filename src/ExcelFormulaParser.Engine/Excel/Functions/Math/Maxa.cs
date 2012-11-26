using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.Math
{
    public class Maxa : ExcelFunction
    {
        private readonly DoubleEnumerableArgConverter _argConverter;

        public Maxa()
            : this(new DoubleEnumerableArgConverter())
        {

        }

        public Maxa(DoubleEnumerableArgConverter argConverter)
        {
            Require.That(argConverter).Named("argConverter").IsNotNull();
            _argConverter = argConverter;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var values = _argConverter.ConvertArgsIncludingOtherTypes(arguments);
            return CreateResult(values.Max(), DataType.Decimal);
        }
    }
}
