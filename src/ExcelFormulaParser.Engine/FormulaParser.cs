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
        public FormulaParser(ExcelDataProvider excelDataProvider)
        {
            Configure(x =>
            {
                x.SetLexer(new Lexer())
                    .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, x.FunctionRepository))
                    .SetExpresionCompiler(new ExpressionCompiler());
                x.FunctionRepository.LoadModule(new BuiltInFunctions());
            });
        }

        public void Configure(Action<ParsingConfiguration> configMethod)
        {
            var config = ParsingConfiguration.Create();
            configMethod.Invoke(config);
            _lexer = config.Lexer;
            _graphBuilder = config.GraphBuilder;
            _compiler = config.ExpressionCompiler;
        }

        private ILexer _lexer;
        private IExpressionGraphBuilder _graphBuilder;
        private IExpressionCompiler _compiler;

        public object Parse(string formula)
        {
            var tokens = _lexer.Tokenize(formula);
            var graph = _graphBuilder.Build(tokens);
            return _compiler.Compile(graph.Expressions).Result;
        }
    }
}
