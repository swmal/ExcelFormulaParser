using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine
{
    public class ExcelDataItem
    {
        public ExcelDataItem(object val, int colIndex, int rowIndex)
        {
            Value = val;
            ColIndex = colIndex;
            RowIndex = rowIndex;
        }

        public int ColIndex { get; private set; }

        public int RowIndex { get; private set; }

        public object Value { get; private set; }
    }
}
