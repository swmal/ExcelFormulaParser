using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;
using Rhino.Mocks;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExGraph = ExcelFormulaParser.Engine.ExpressionGraph.ExpressionGraph;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class FormulaParserTests
    {
        private FormulaParser _parser;

        [TestInitialize]
        public void Setup()
        {
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(provider);

        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void ParserShouldCallLexer()
        {
            var lexer = MockRepository.GenerateStub<ILexer>();
            lexer.Stub(x => x.Tokenize("ABC")).Return(Enumerable.Empty<Token>());
            _parser.Configure(x => x.SetLexer(lexer));

            _parser.Parse("ABC");

            lexer.AssertWasCalled(x => x.Tokenize("ABC"));
        }

        [TestMethod]
        public void ParserShouldCallGraphBuilder()
        {
            var lexer = MockRepository.GenerateStub<ILexer>();
            var tokens = new List<Token>();
            lexer.Stub(x => x.Tokenize("ABC")).Return(tokens);
            var graphBuilder = MockRepository.GenerateStub<IExpressionGraphBuilder>();
            graphBuilder.Stub(x => x.Build(tokens)).Return(new ExGraph());

            _parser.Configure(config =>
                {
                    config
                        .SetLexer(lexer)
                        .SetGraphBuilder(graphBuilder);
                });

            _parser.Parse("ABC");

            graphBuilder.AssertWasCalled(x => x.Build(tokens));
        }

        [TestMethod]
        public void ParserShouldCallCompiler()
        {
            var lexer = MockRepository.GenerateStub<ILexer>();
            var tokens = new List<Token>();
            lexer.Stub(x => x.Tokenize("ABC")).Return(tokens);
            var expectedGraph = new ExGraph();
            expectedGraph.Add(new StringExpression("asdf"));
            var graphBuilder = MockRepository.GenerateStub<IExpressionGraphBuilder>();
            graphBuilder.Stub(x => x.Build(tokens)).Return(expectedGraph);
            var compiler = MockRepository.GenerateStub<IExpressionCompiler>();
            compiler.Stub(x => x.Compile(expectedGraph.Expressions)).Return(new CompileResult(0, DataType.Integer));

            _parser.Configure(config =>
            {
                config
                    .SetLexer(lexer)
                    .SetGraphBuilder(graphBuilder)
                    .SetExpresionCompiler(compiler);
            });

            _parser.Parse("ABC");

            compiler.AssertWasCalled(x => x.Compile(expectedGraph.Expressions));
        }
    }
}
