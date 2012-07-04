using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public interface IFunctionModule
    {
        IDictionary<string, ExcelFunction> Functions { get; }
    }
}
