using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine
{
    public class ParsingContext
    {
        private ParsingContext() { }

        public FormulaParser Parser { get; set; }

        public ParsingConfiguration Configuration { get; set; }

        public Ranges Ranges { get; private set; }

        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            context.Ranges = new Ranges();
            return context;
        }
    }
}
