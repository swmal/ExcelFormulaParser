using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using System.Diagnostics;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalculationChainBuilder
    {
        private readonly ParsingContext _parsingContext;

        public CalculationChainBuilder(ParsingContext parsingContext)
        {
            Require.That(parsingContext).Named("parsingContext").IsNotNull();
            _parsingContext = parsingContext;
        }

        public virtual CalculationChain Build()
        {
            return Build("Test");
        }

        public virtual CalculationChain Build(string worksheetName)
        {
            var graphs = new DependencyGraphs();
            var lexer = _parsingContext.Configuration.Lexer;
            var cells = string.IsNullOrEmpty(worksheetName) ?
                _parsingContext.ExcelDataProvider.GetWorkbookFormulas() :
                _parsingContext.ExcelDataProvider.GetWorksheetFormulas(worksheetName);
            var sw = new Stopwatch();
            long part1 = 0, part2 = 0;
            foreach (var formulaCell in cells) 
            {
                sw.Start();
                var currentCell = formulaCell.Key;
                var formula = formulaCell.Value;
                var r = lexer.Tokenize(formula);
                var toAddresses = r.Where(x => x.TokenType == TokenType.ExcelAddress);
                sw.Stop();
                part1 += sw.ElapsedMilliseconds;
                sw.Reset();
                sw.Start();
                foreach (var address in toAddresses)
                {
                    if (cells.ContainsKey(address.Value))
                    {
                        graphs.Add(currentCell, address.Value);
                    }
                }
                sw.Stop();
                part2 += sw.ElapsedMilliseconds;
                sw.Reset();
            }
            var calcChain = new CalculationChain();
            foreach (var graph in graphs.GetGraphs())
            {
                foreach (var formula in graph.OrderedFormulas)
                {
                    calcChain.Add(new CalculationCell { Address = formula });
                }
            }
            return calcChain;
        }
    }
}
