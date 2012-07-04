using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine
{
    public class ParsingContext
    {
        private Dictionary<string, bool> _referencedCells = new Dictionary<string, bool>();

        public ParsingConfiguration Configuration { get; set; }

        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            return context;
        }

        public bool IsReferencedCell(string cellAddress)
        {
            if (string.IsNullOrEmpty(cellAddress)) return false;
            return _referencedCells.ContainsKey(cellAddress);
        }

        public void AddReferencedCell(string cellAddress)
        {
            if (!_referencedCells.ContainsKey(cellAddress))
            {
                _referencedCells.Add(cellAddress, true);
            }
        }
    }
}
