using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public abstract class LookupFunction : ExcelFunction
    {
        public override bool IsLookupFuction
        {
            get
            {
                return true;
            }
        }
    }
}
