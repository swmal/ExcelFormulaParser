﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public class FunctionArgument
    {
        public FunctionArgument(object val)
        {
            Value = val;
        }

        private ExcelCellState _excelCellState;

        public void SetExcelStateFlag(ExcelCellState state)
        {
            _excelCellState |= state;
        }

        public bool ExcelStateFlagIsSet(ExcelCellState state)
        {
            return (_excelCellState & state) != 0;
        }

        public object Value { get; private set; }

        public Type Type
        {
            get { return Value != null ? Value.GetType() : null; }
        }
    }
}
