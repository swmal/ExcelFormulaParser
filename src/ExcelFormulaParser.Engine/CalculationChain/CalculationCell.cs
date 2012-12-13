using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalculationCell
    {
        public string Address
        {
            get;
            set;
        }

        public string Worksheet
        {
            get;
            set;
        }
    }
}
