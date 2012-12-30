using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalcCellRepository
    {
        private readonly Dictionary<object, CalcCell> _cells = new Dictionary<object,CalcCell>();
        private readonly Dictionary<object, List<object>> _addressIndex = new Dictionary<object, List<object>>();

        private readonly CalcChainContext _context;

        private CalcCellRepository(CalcChainContext context)
        {
            _context = context;
        }

        public static CalcCellRepository Create(CalcChainContext context)
        {
            return new CalcCellRepository(context);
        }

        public virtual CalcCell AddOrGet(string address)
        {
            var existingCell = GetCell(address);
            return existingCell ?? CreateAndAddNewCell(address);
        }

        private CalcCell CreateAndAddNewCell(string address)
        {
            var newCell = CalcCell.Create(address, _context);
            _cells.Add(newCell.Id, newCell);
            var hashCode = GetHashCode(address);
            if (!_addressIndex.ContainsKey(hashCode))
            {
                _addressIndex[hashCode] = new List<object>();
            }
            _addressIndex[hashCode].Add(newCell.Id);
            return newCell;
        }

        public virtual CalcCell GetCell(object id)
        {
            if (_cells.ContainsKey(id))
            {
                return _cells[id];
            }
            return null;
        }

        public virtual CalcCell GetCell(string address)
        {
            CalcCell result = null;
            var loweredAddress = address.ToUpperInvariant();
            var hashCode = GetHashCode(address);
            if (!_addressIndex.ContainsKey(hashCode)) return null;
            foreach (var id in _addressIndex[hashCode])
            {
                var candidate = _cells[id];
                if (candidate.Address.ToUpperInvariant().Equals(loweredAddress))
                {
                    result = candidate;
                    break;
                }
            }
            return result;
        }

        private static int GetHashCode(string input)
        {
            return input.ToUpperInvariant().GetHashCode();
        }
    }
}
