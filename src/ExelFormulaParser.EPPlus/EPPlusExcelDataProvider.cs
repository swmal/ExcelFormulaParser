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

        public EPPlusExcelDataProvider(ExcelPackage package)
        {
            _package = package;
        }

        public override IEnumerable<ExcelDataItem> GetRangeValues(string address)
        {
            var returnList = new List<ExcelDataItem>();
            var startAddress = RangeAddress.Parse(address);
            if (AddressHasWorkbookName(address))
            {
                _currentWorksheet = _package.Workbook.Worksheets[GetWorksheetName(address)];
            }
            else if(_currentWorksheet == null)
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
            var range = _currentWorksheet.Cells[GetRangeAddress(address)];
            if (range.Value is object[,])
            {
                var arr = (object[,])range.Value;
                var nRows = arr.GetUpperBound(0);
                var nCols = arr.GetUpperBound(1);
                for(int row = 0; row <= nRows; row++)
                {
                    for (int col = 0; col <= nCols; col++)
                    {
                        returnList.Add(new ExcelDataItem(arr[row, col], startAddress.FromCol + col, startAddress.FromRow + row));
                    }

                }
            }
            else 
            { 
                returnList.Add(new ExcelDataItem(range.Value, startAddress.FromCol, startAddress.FromRow)); 
            }
            return returnList;
        }

        private static bool AddressHasWorkbookName(string address)
        {
            return address.Contains("!");
        }

        private static string GetWorksheetName(string address)
        {
            return address.Split('!').First();
        }

        private static string GetRangeAddress(string address)
        {
            if (address.Contains("!"))
            {
                return address.Split('!')[1];
            }
            return address;
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
