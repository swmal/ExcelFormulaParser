using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public abstract class ArgumentParser
    {
        public abstract object Parse(object obj);
    }
}
