using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine;
using OfficeOpenXml;
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

        public override IEnumerable<ExcelDataItem> GetRangeValues(string address)
        {
            var returnList = new List<ExcelDataItem>();
            var startAddress = _rangeAddressFactory.Create(address);
            var addressInfo = ExcelAddressInfo.Parse(address);
            if (addressInfo.WorksheetIsSpecified)
            {
                _currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
            }
            else if(_currentWorksheet == null)
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
            var range = _currentWorksheet.Cells[addressInfo.AddressOnSheet];
            foreach (var cell in range)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    //parse formula
                    returnList.Add(new ExcelDataItem("=" + cell.Formula, cell.Start.Column, cell.Start.Row));
                }
                else
                {
                    returnList.Add(new ExcelDataItem(cell.Value, cell.Start.Column, cell.Start.Row));
                }

            }
            //if (range.Value is object[,])
            //{
            //    var arr = (object[,])range.Value;
            //    var nRows = arr.GetUpperBound(0);
            //    var nCols = arr.GetUpperBound(1);
            //    for(int row = 0; row <= nRows; row++)
            //    {
            //        for (int col = 0; col <= nCols; col++)
            //        {
            //            returnList.Add(new ExcelDataItem(arr[row, col], startAddress.FromCol + col, startAddress.FromRow + row));
            //        }

            //    }
            //}
            //else 
            //{ 
            //    returnList.Add(new ExcelDataItem(range.Value, startAddress.FromCol, startAddress.FromRow)); 
            //}
            return returnList;
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
