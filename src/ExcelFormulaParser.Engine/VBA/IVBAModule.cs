using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.VBA.Functions
{
    public interface IVBAModule
    {
        IDictionary<string, VBAFunction> Functions { get; }
    }
}
