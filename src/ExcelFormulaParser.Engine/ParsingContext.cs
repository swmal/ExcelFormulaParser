using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine
{
    public class ParsingContext : IParsingLifetimeEventHandler
    {
        private ParsingContext() { }

        public FormulaParser Parser { get; set; }

        public ParsingConfiguration Configuration { get; set; }

        public FormulaDependencies Dependencies { get; private set; }

        public ParsingScopes Scopes { get; private set; }

        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            context.Scopes = new ParsingScopes(context);
            context.Dependencies = new FormulaDependencies();
            return context;
        }

        void IParsingLifetimeEventHandler.ParsingCompleted()
        {
            Dependencies.Clear();
        }
    }
}
