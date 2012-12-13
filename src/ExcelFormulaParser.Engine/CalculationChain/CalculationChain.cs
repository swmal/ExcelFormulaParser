using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalculationChain
    {
        private List<CalculationCell> _calcCells = new List<CalculationCell>();

        public void Add(CalculationCell cell)
        {
            _calcCells.Add(cell);
        }

        public IEnumerable<CalculationCell> Cells
        {
            get { return _calcCells; }
        }
    }
}
