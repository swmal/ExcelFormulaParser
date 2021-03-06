﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.EPPlus
{
    public class EPPlusExcelDataProvider : ExcelDataProvider
    {
        private readonly ExcelPackage _package;
        private ExcelWorksheet _currentWorksheet;
        private RangeAddressFactory _rangeAddressFactory;

        public EPPlusExcelDataProvider(ExcelPackage package)
        {
            _package = package;
            _rangeAddressFactory = new RangeAddressFactory(this);
        }

        public override IEnumerable<string> GetWorksheetNames()
        {
            var names = new List<string>();
            foreach(var ws in _package.Workbook.Worksheets)
            {
                names.Add(ws.Name);
            }
            return names;
        }

        public override IDictionary<string, string> GetWorksheetFormulas(string sheetName)
        {
            var ws = _package.Workbook.Worksheets[sheetName];
            return ((ICalcEngineFormulaInfo)ws).GetFormulas();
        }

        public override IDictionary<string, string> GetWorkbookFormulas()
        {
            var wb = (ICalcEngineFormulaInfo)_package.Workbook;
            return wb.GetFormulas();
        }

        public override IDictionary<string, object> GetWorkbookNameValues()
        {
            var wb = (ICalcEngineFormulaInfo)_package.Workbook;
            return wb.GetNameValues();
        }

        public override IEnumerable<ExcelCell> GetRangeValues(string address)
        {
            var returnList = new List<ExcelCell>();
            var addressInfo = ExcelAddressInfo.Parse(address);
            SetCurrentWorksheet(addressInfo);
            var range = _currentWorksheet.Cells[addressInfo.AddressOnSheet];
            foreach (var cell in range)
            {
                returnList.Add(new ExcelCell(cell.Value, cell.Formula, cell.Start.Column, cell.Start.Row));
            }
            return returnList;
        }

        public override ExcelCell GetCellValue(string address)
        {
            var addressInfo = ExcelAddressInfo.Parse(address);
            SetCurrentWorksheet(addressInfo);
            var cell = _currentWorksheet.Cells[addressInfo.AddressOnSheet].FirstOrDefault();
            if (cell != null)
            {
                return new ExcelCell(cell.Value, cell.Formula, cell.Start.Column, cell.Start.Row);
            }
            return null;
        }

        public override ExcelCell GetCellValue(int row, int col)
        {
            var cell = _currentWorksheet.Cells[row, col];
            if (cell != null)
            {
                return new ExcelCell(cell.Value, cell.Formula, cell.Start.Column, cell.Start.Row);
            }
            return null;
        }

        private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
        {
            if (addressInfo.WorksheetIsSpecified)
            {
                _currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
            }
            else if (_currentWorksheet == null)
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
        }

        public override void SetCellValue(string address, object value)
        {
            var addressInfo = ExcelAddressInfo.Parse(address);
            var ra = _rangeAddressFactory.Create(address);
            SetCurrentWorksheet(addressInfo);
            var valueInfo = (ICalcEngineValueInfo)_currentWorksheet;
            valueInfo.SetFormulaValue(ra.FromRow + 1, ra.FromCol + 1, value);
        }

        public override void Dispose()
        {
            _package.Dispose();
        }

        public override int ExcelMaxColumns
        {
            get { return ExcelPackage.MaxColumns; }
        }

        public override int ExcelMaxRows
        {
            get { return ExcelPackage.MaxRows; }
        }
    }
}
