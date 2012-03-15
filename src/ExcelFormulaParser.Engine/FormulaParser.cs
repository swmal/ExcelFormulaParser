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
            :this(new Lexer(), new ExpressionGraphBuilder(excelDataProvider), new ExpressionCompiler())
        {

        }

        public FormulaParser(ILexer lexer, IExpressionGraphBuilder graphBuilder, IExpressionCompiler compiler)
        {
            _lexer = lexer;
            _graphBuilder = graphBuilder;
            _compiler = compiler;
            FunctionRepository.Clear();
            FunctionRepository.LoadModule(new BuiltInFunctions());
        }

        private readonly ILexer _lexer;
        private readonly IExpressionGraphBuilder _graphBuilder;
        private readonly IExpressionCompiler _compiler;

        public object Parse(string formula)
        {
            var tokens = _lexer.Tokenize(formula);
            var graph = _graphBuilder.Build(tokens);
            return _compiler.Compile(graph.Expressions).Result;
        }
    }
}
