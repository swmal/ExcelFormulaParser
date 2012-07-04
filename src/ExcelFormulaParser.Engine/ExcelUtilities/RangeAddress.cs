using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class RangeAddress
    {
        private static readonly AddressTranslator _addressTranslator = new AddressTranslator();

        public string Worksheet { get; private set; }

        public int FromCol { get; private set; }

        public int ToCol { get; private set; }

        public int FromRow { get; private set; }

        public int ToRow { get; private set; }

        /// <summary>
        /// Creates a <see cref="RangeAddress"/> instance out of a valid
        /// speadsheet address, ex. A1, A1:B2, Worksheet!A1:B2
        /// </summary>
        /// <param name="range">The range to parse</param>
        /// <returns></returns>
        public static RangeAddress Parse(string range)
        {
            var worksheet = string.Empty;
            var rangeAddress = range;
            if (range.Contains("!"))
            {
                worksheet = range.Split('!')[0];
                rangeAddress = range.Split('!')[1];
            }
            if (!rangeAddress.Contains(":"))
            {
                return HandleSingleCellAddress(rangeAddress, worksheet);
            }
            return HandleMultipleCellAddress(rangeAddress, worksheet);
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

        private static RangeAddress HandleSingleCellAddress(string range, string worksheet)
        {
            int col, row;
            _addressTranslator.ToColAndRow(range, out col, out row);
            return new RangeAddress { FromCol = col, ToCol = col, FromRow = row, ToRow = row, Worksheet = worksheet };
        }

        private static RangeAddress HandleMultipleCellAddress(string range, string worksheet)
        {
            var rangeArr = range.Split(':');
            int fromCol, fromRow;
            _addressTranslator.ToColAndRow(rangeArr[0], out fromCol, out fromRow);
            int toCol, toRow;
            _addressTranslator.ToColAndRow(rangeArr[1], out toCol, out toRow);
            return new RangeAddress { FromCol = fromCol, ToCol = toCol, FromRow = fromRow, ToRow = toRow, Worksheet = worksheet };
        }
    }
}
