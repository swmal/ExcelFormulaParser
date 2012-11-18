using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Engine.Excel.Functions.RefAndLookup
{
    public class LookupNavigator
    {
        private readonly LookupDirection _direction;
        private readonly LookupArguments _arguments;
        private readonly ExcelDataProvider _excelDataProvider;
        private RangeAddress _rangeAddress;
        private int _currentRow;
        private int _currentCol;

        public LookupNavigator(LookupDirection direction, LookupArguments arguments, ExcelDataProvider excelDataProvider)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _direction = direction;
            _arguments = arguments;
            _excelDataProvider = excelDataProvider;
            Initialize();
        }

        private void Initialize()
        {
            var factory = new RangeAddressFactory(_excelDataProvider);
            _rangeAddress = factory.Create(_arguments.RangeAddress);
            _currentCol = _rangeAddress.FromCol;
            _currentRow = _rangeAddress.FromRow;
            SetCurrentValue();
        }

        private void SetCurrentValue()
        {
            var cellValue = _excelDataProvider.GetCellValue(_currentRow, _currentCol);
            CurrentValue = cellValue != null ? cellValue.Value : null;
        }

        private bool HasNext()
        {
            if (_direction == LookupDirection.Vertical)
            {
                return _currentRow < _rangeAddress.ToRow;
            }
            else
            {
                return _currentCol < _rangeAddress.ToCol;
            }
        }

        public bool MoveNext()
        {
            if (!HasNext()) return false;
            if (_direction == LookupDirection.Vertical)
            {
                _currentRow++;
            }
            else
            {
                _currentCol++;
            }
            SetCurrentValue();
            return true;
        }

        public object CurrentValue
        {
            get;
            private set;
        }

        public object GetLookupValue()
        {
            var row = _currentRow;
            var col = _currentCol;
            if (_direction == LookupDirection.Vertical)
            {
                col += _arguments.LookupIndex - 1;
            }
            else
            {
                row += _arguments.LookupIndex - 1;
            }
            var cellValue = _excelDataProvider.GetCellValue(row, col);
            return cellValue != null ? cellValue.Value : null;
        }
    }
}
