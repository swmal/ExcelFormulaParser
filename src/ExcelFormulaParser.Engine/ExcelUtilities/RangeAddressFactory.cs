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
        private readonly AddressTranslator _addressTranslator;
        private readonly IndexToAddressTranslator _indexToAddressTranslator;

        public RangeAddressFactory(ExcelDataProvider excelDataProvider)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _addressTranslator = new AddressTranslator(excelDataProvider);
            _indexToAddressTranslator = new IndexToAddressTranslator(excelDataProvider, ExcelReferenceType.RelativeRowAndColumn);
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
            var addressInfo = ExcelAddressInfo.Parse(range);
            var rangeAddress = new RangeAddress()
            {
                Address = range,
                Worksheet = addressInfo.Worksheet
            };
           
            if (addressInfo.IsMultipleCells)
            {
                HandleMultipleCellAddress(rangeAddress, addressInfo);
            }
            else
            {
                HandleSingleCellAddress(rangeAddress, addressInfo);
            }
            return rangeAddress;
        }

        private void HandleSingleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
        {
            int col, row;
            _addressTranslator.ToColAndRow(addressInfo.StartCell, out col, out row);
            rangeAddress.FromCol = col;
            rangeAddress.ToCol = col;
            rangeAddress.FromRow = row;
            rangeAddress.ToRow = row;
        }

        private void HandleMultipleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
        {
            int fromCol, fromRow;
            _addressTranslator.ToColAndRow(addressInfo.StartCell, out fromCol, out fromRow);
            int toCol, toRow;
            _addressTranslator.ToColAndRow(addressInfo.EndCell, out toCol, out toRow, AddressTranslator.RangeCalculationBehaviour.LastPart);
            rangeAddress.FromCol = fromCol;
            rangeAddress.ToCol = toCol;
            rangeAddress.FromRow = fromRow;
            rangeAddress.ToRow = toRow;
        }
    }
}
