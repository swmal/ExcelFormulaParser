using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Excel.Functions;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class Lexer : ILexer
    {
        public Lexer(FunctionRepository functionRepository, NameValueProvider nameValueProvider)
            :this(new SourceCodeTokenizer(functionRepository, nameValueProvider), new SyntacticAnalyzer())
        {

        }

        public Lexer(ISourceCodeTokenizer tokenizer, ISyntacticAnalyzer analyzer)
        {
            _tokenizer = tokenizer;
            _analyzer = analyzer;
        }

        private readonly ISourceCodeTokenizer _tokenizer;
        private readonly ISyntacticAnalyzer _analyzer;

        public IEnumerable<Token> Tokenize(string input)
        {
            var tokens = _tokenizer.Tokenize(input);
            _analyzer.Analyze(tokens);
            return tokens;
        }
    }
}
