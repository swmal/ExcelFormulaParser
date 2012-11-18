using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class LookupArguments
    {
        public LookupArguments(IEnumerable<FunctionArgument> arguments)
            : this(arguments, new ArgumentParsers())
        {

        }

        public LookupArguments(IEnumerable<FunctionArgument> arguments, ArgumentParsers argumentParsers)
        {
            _argumentParsers = argumentParsers;
            SearchedValue = arguments.ElementAt(0).Value;
            RangeAddress = arguments.ElementAt(1).Value.ToString();
            LookupIndex = (int)_argumentParsers.GetParser(DataType.Integer).Parse(arguments.ElementAt(2).Value);
            if (arguments.Count() > 3)
            {
                RangeLookup = (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(arguments.ElementAt(3).Value);
            }
        }

        private readonly ArgumentParsers _argumentParsers;

        public object SearchedValue { get; private set; }

        public string RangeAddress { get; private set; }

        public int LookupIndex { get; private set; }

        public bool RangeLookup { get; private set; }
    }
}
