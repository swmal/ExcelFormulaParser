using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class RangeAddressFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private static readonly AddressTranslator _addressTranslator = new AddressTranslator();
        private static readonly IndexToAddressTranslator _indexToAddressTranslator = new IndexToAddressTranslator();

        public RangeAddressFactory(ExcelDataProvider excelDataProvider)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
        }

        public RangeAddress Create(int col, int row)
        {
            return new RangeAddress()
            {
                Address = _indexToAddressTranslator.ToAddress(col, row),
                Worksheet = string.Empty,
                FromCol = col,
                ToCol = col,
                FromRow = row,
                ToRow = row
            };
        }

        public RangeAddress Create(string range)
        {
            Require.That(range).Named("range").IsNotNullOrEmpty();
            var worksheet = string.Empty;
            var worksheetAddress = range;
            if (range.Contains("!"))
            {
                worksheet = range.Split('!')[0];
                worksheetAddress = range.Split('!')[1];
            }
            var rangeAddress = new RangeAddress
            {
                Address = range,
                Worksheet = worksheet
            };
            if (!worksheetAddress.Contains(":"))
            {
                HandleSingleCellAddress(rangeAddress, worksheetAddress);
            }
            else
            {
                HandleMultipleCellAddress(rangeAddress, worksheetAddress);
            }
            return rangeAddress;
        }

        private static void HandleSingleCellAddress(RangeAddress rangeAddress, string range)
        {
            int col, row;
            _addressTranslator.ToColAndRow(range, out col, out row);
            rangeAddress.FromCol = col;
            rangeAddress.ToCol = col;
            rangeAddress.FromRow = row;
            rangeAddress.ToRow = row;
        }

        private static void HandleMultipleCellAddress(RangeAddress rangeAddress, string range)
        {
            var rangeArr = range.Split(':');
            int fromCol, fromRow;
            _addressTranslator.ToColAndRow(rangeArr[0], out fromCol, out fromRow);
            int toCol, toRow;
            _addressTranslator.ToColAndRow(rangeArr[1], out toCol, out toRow);
            rangeAddress.FromCol = fromCol;
            rangeAddress.ToCol = toCol;
            rangeAddress.FromRow = fromRow;
            rangeAddress.ToRow = toRow;
        }
    }
}
