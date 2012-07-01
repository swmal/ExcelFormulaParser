﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine;
using OfficeOpenXml;

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

        public override IEnumerable<object> GetRangeValues(string address)
        {
            var returnList = new List<object>();
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
                foreach (var obj in (object[,])range.Value)
                {
                    returnList.Add(obj);
                }
            }
            else 
            { 
                returnList.Add(range.Value); 
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
    }
}