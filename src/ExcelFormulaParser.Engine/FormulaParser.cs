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

namespace ExcelFormulaParser.Engine
{
    public class FormulaParser
    {
        private readonly ParsingContext _context = ParsingContext.Create();

        public FormulaParser(ExcelDataProvider excelDataProvider)
        {
            Configure(x =>
            {
                x.SetLexer(new Lexer(_context.Configuration.FunctionRepository))
                    .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, _context))
                    .SetExpresionCompiler(new ExpressionCompiler());
                x.FunctionRepository.LoadModule(new BuiltInFunctions());
            });
        }

        public void Configure(Action<ParsingConfiguration> configMethod)
        {
            configMethod.Invoke(_context.Configuration);
            _lexer = _context.Configuration.Lexer ?? _lexer;
            _graphBuilder = _context.Configuration.GraphBuilder ?? _graphBuilder;
            _compiler = _context.Configuration.ExpressionCompiler ?? _compiler;
        }

        private ILexer _lexer;
        private IExpressionGraphBuilder _graphBuilder;
        private IExpressionCompiler _compiler;

        public virtual object Parse(string formula)
        {
            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
            {
                return formula;
            }
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
}
