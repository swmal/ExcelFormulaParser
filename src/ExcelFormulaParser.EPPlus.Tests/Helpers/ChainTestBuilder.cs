using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace ExcelFormulaParser.EPPlus.Tests.Helpers
{
    public class ChainTestBuilder
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        public ChainTestBuilder()
        {
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("Test");
        }

        public ExcelPackage Build()
        {
            _worksheet.Cells["B1"].Formula = "SUM(B2:B1000)";
            _worksheet.Cells["C1"].Formula = "SUM(C2:C15001)";

            for (var x = 2; x < 15000; x++)
            {
                _worksheet.Cells[x, 1].Value = x - 1;
                if (x == 2)
                {
                    _worksheet.Cells[x, 2].Formula = "A" + x.ToString();
                }
                else
                {
                    _worksheet.Cells[x, 2].Formula = "A" + x.ToString() + " + B" + (x - 1).ToString();
                }
                _worksheet.Cells[x, 3].Formula = "C" + (x + 1).ToString() + " + A" + x.ToString();
            }

            return _package;
        }

    }
}
