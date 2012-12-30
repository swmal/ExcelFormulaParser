using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalcChainContextBuilder
    {
        private ParsingContext _parsingContext;

        public CalcChainContext Build(ParsingContext context, string worksheetName)
        {
            _parsingContext = context;
            var chainContext = CalcChainContext.Create(context.Configuration.IdProvider);
            var lexer = context.Configuration.Lexer;
            var cells = string.IsNullOrEmpty(worksheetName) ?
                _parsingContext.ExcelDataProvider.GetWorkbookFormulas() :
                _parsingContext.ExcelDataProvider.GetWorksheetFormulas(worksheetName);
            foreach (var cellInfo in cells)
            {
                var address = cellInfo.Key;
                var cell = chainContext.CalcCells.AddOrGet(address);
                var formula = cellInfo.Value;
                var r = lexer.Tokenize(formula);
                var toAddresses = r.Where(x => x.TokenType == TokenType.ExcelAddress);
                foreach (var toAddress in toAddresses)
                {
                    var rangeAddress = context.RangeAddressFactory.Create(toAddress.Value);
                    var rangeCells = new List<string>();
                    if (rangeAddress.FromRow < rangeAddress.ToRow || rangeAddress.FromCol < rangeAddress.ToCol)
                    {
                        for (var col = rangeAddress.FromCol; col <= rangeAddress.ToCol; col++)
                        {
                            for (var row = rangeAddress.FromRow; row <= rangeAddress.ToRow; row++)
                            {
                                rangeCells.Add(context.RangeAddressFactory.Create(col + 1, row + 1).Address);
                            }
                        }
                    }
                    else
                    {
                        rangeCells.Add(toAddress.Value);
                    }
                    CalcChain chain;
                    foreach (var rangeCell in rangeCells)
                    {
                        var toCell = chainContext.CalcCells.AddOrGet(rangeCell);
                        var chainId = cell.GetCalcChainId();
                        if (chainId == null)
                        {
                            chain = CalcChain.Create(context.Configuration.IdProvider);
                            chainContext.AddCalcChain(chain);
                        }
                        else
                        {
                            chain = chainContext.CalcChains.Where(x => x.Id == chainId).FirstOrDefault();
                        }
                        chain.Add(toCell);
                        cell.AddRelationTo(toCell, chain);
                        toCell.AddRelationFrom(cell, chain);
                    }
                }
            }

            return chainContext;

        }
    }
}
