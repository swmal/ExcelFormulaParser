using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine
{
    public class ParsingContext
    {
        public ParsingConfiguration Configuration { get; set; }

        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            return context;
        }
    }
}
