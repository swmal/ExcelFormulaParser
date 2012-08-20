using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.EPPlus
{
    public static class ExtensionMethods
    {
        public static void Calculate(this ExcelWorksheet worksheet, ExcelPackage package)
        {
            var provider = new EPPlusExcelDataProvider(package);
            var parser = new FormulaParser(provider);
            ParseFormulas(provider, parser, worksheet);
        }

        public static void Calculate(this ExcelWorkbook workbook, ExcelPackage package)
        {
            var provider = new EPPlusExcelDataProvider(package);
            var parser = new FormulaParser(provider);
            foreach (var worksheet in workbook.Worksheets)
            {
                ParseFormulas(provider, parser, worksheet);
            }
        }

        private static void ParseFormulas(EPPlusExcelDataProvider provider, FormulaParser parser, ExcelWorksheet worksheet)
        {
            var formulas = provider.GetWorksheetFormulas(worksheet.Name);
            foreach (var formulaAddress in formulas.Keys)
            {
                var result = parser.ParseAt(formulaAddress);
                provider.SetCellValue(formulaAddress, result);
            }
        }
    }
}
