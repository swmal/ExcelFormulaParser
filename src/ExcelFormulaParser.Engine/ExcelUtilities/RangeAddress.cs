using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class RangeAddress
    {
        private static readonly AddressTranslator _addressTranslator = new AddressTranslator();
        private static readonly IndexToAddressTranslator _indexToAddressTranslator = new IndexToAddressTranslator();

        private string _rangeAddress = string.Empty;

        public string Worksheet { get; private set; }

        public int FromCol { get; private set; }

        public int ToCol { get; private set; }

        public int FromRow { get; private set; }

        public int ToRow { get; private set; }

        public override string ToString()
        {
            return _rangeAddress;
        }

        /// <summary>
        /// Creates a <see cref="RangeAddress"/> instance out of a valid
        /// speadsheet address, ex. A1, A1:B2, Worksheet!A1:B2
        /// </summary>
        /// <param name="range">The range to parse</param>
        /// <returns></returns>
        public static RangeAddress Parse(string range)
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
                _rangeAddress = range,
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

        public static RangeAddress Create(int col, int row)
        {
            return new RangeAddress()
            {
                _rangeAddress = _indexToAddressTranslator.ToAddress(col, row),
                Worksheet = string.Empty,
                FromCol = col,
                ToCol = col,
                FromRow = row,
                ToRow = row
            };
        }

        private static RangeAddress _empty = new RangeAddress();
        public static RangeAddress Empty
        {
            get { return _empty; }
        }

        /// <summary>
        /// Returns true if this range collides (full or partly) with the supplied range
        /// </summary>
        /// <param name="other">The range to check</param>
        /// <returns></returns>
        public bool CollidesWith(RangeAddress other)
        {
            if (other.Worksheet != Worksheet)
            {
                return false;
            }
            if (other.FromRow > ToRow || other.FromCol > ToCol
                ||
                FromRow > other.ToRow || FromCol > other.ToCol)
            {
                return false;
            }
            return true;
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
