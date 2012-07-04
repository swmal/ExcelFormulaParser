using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine;
using ExcelFormulaParser.Engine.Excel.Operators;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.Excel;
using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine
{
    public class FormulaParser
    {
        private readonly ParsingContext _parsingContext;
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly RangeAddressFactory _rangeAddressFactory;

        public FormulaParser(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ParsingContext.Create(), new RangeAddressFactory(excelDataProvider))
        {
           
        }

        public FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext, RangeAddressFactory rangeAddressFactory)
        {
            parsingContext.Parser = this;
            _parsingContext = parsingContext;
            _excelDataProvider = excelDataProvider;
            _rangeAddressFactory = rangeAddressFactory;
            Configure(configuration =>
            {
                configuration
                    .SetLexer(new Lexer(_parsingContext.Configuration.FunctionRepository))
                    .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, _parsingContext))
                    .SetExpresionCompiler(new ExpressionCompiler())
                    .FunctionRepository.LoadModule(new BuiltInFunctions());
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
                _parsingContext.Dependencies.AddFormulaScope(scope);
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
            return Parse(formula, _rangeAddressFactory.Create(address));
        }

        public virtual object Parse(string formula)
        {
            return Parse(formula, RangeAddress.Empty);
        }

        public virtual object ParseAt(string address)
        {
            Require.That(address).Named("address").IsNotNullOrEmpty();
            var dataItem = _excelDataProvider.GetRangeValues(address).First();
            if (dataItem.Value == null && dataItem.Formula == null) return null;
            if (!string.IsNullOrEmpty(dataItem.Formula))
            {
                return Parse(dataItem.Formula, _rangeAddressFactory.Create(address));
            }
            return Parse(dataItem.Value.ToString(), _rangeAddressFactory.Create(address));
        }
    }
}
