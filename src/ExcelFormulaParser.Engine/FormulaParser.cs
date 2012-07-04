using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Engine.VBA.Operators;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.VBA;
using ExcelFormulaParser.Engine.VBA.Functions;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine
{
    public class FormulaParser
    {
        private readonly ParsingContext _parsingContext;
        private readonly ExcelDataProvider _excelDataProvider;

        public FormulaParser(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ParsingContext.Create())
        {
           
        }

        public FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
        {
            parsingContext.Parser = this;
            _parsingContext = parsingContext;
            _excelDataProvider = excelDataProvider;
            Configure(x =>
            {
                x.SetLexer(new Lexer(_parsingContext.Configuration.FunctionRepository))
                    .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, _parsingContext))
                    .SetExpresionCompiler(new ExpressionCompiler());
                x.FunctionRepository.LoadModule(new BuiltInFunctions());
            }); 
        }

        public void Configure(Action<ParsingConfiguration> configMethod)
        {
            configMethod.Invoke(_parsingContext.Configuration);
            _lexer = _parsingContext.Configuration.Lexer ?? _lexer;
            _graphBuilder = _parsingContext.Configuration.GraphBuilder ?? _graphBuilder;
            _compiler = _parsingContext.Configuration.ExpressionCompiler ?? _compiler;
        }

        private ILexer _lexer;
        private IExpressionGraphBuilder _graphBuilder;
        private IExpressionCompiler _compiler;

        internal virtual object Parse(string formula, RangeAddress rangeAddress)
        {
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                if (!IsFormulaCandidate(formula))
                {
                    return formula;
                }
                _parsingContext.Dependencies.AddFormulaScope(scope);
                formula = formula.Substring(1);
                var tokens = _lexer.Tokenize(formula);
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return null;
                }
                return _compiler.Compile(graph.Expressions).Result;
            }
        }

        public virtual object Parse(string formula, string address)
        {
            return Parse(formula, RangeAddress.Parse(address));
        }

        public virtual object Parse(string formula)
        {
            return Parse(formula, RangeAddress.Empty);
        }

        public virtual object ParseAt(string address)
        {
            var dataItem = _excelDataProvider.GetRangeValues(address).First();
            return Parse(dataItem.Value.ToString(), RangeAddress.Parse(address));
        }

        private static bool IsFormulaCandidate(string formula)
        {
            return !string.IsNullOrEmpty(formula) && formula.StartsWith("=");
        }
    }
}
