using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalcChainContext
    {
        private readonly Dictionary<object, CalcChain> _calcChains = new Dictionary<object, CalcChain>();

        private CalcChainContext(IdProvider idProvider)
        {
            IdProvider = idProvider;
            CalcCells = CalcCellRepository.Create(this);
            
        }

        public static CalcChainContext Create(IdProvider idProvider)
        {
            Require.That(idProvider).Named("idProvider").IsNotNull();
            return new CalcChainContext(idProvider);
        }

        public IdProvider IdProvider { get; private set; }

        public CalcCellRepository CalcCells { get; private set; }

        public IEnumerable<CalcChain> CalcChains
        {
            get { return _calcChains.Values; }
        }

        public virtual void AddCalcChain(CalcChain chain)
        {
            _calcChains.Add(chain.Id, chain);
        }

        public virtual CalcChain GetCalcChain(object id)
        {
            if (!_calcChains.ContainsKey(id)) return null;
            return _calcChains[id];
        }
    }
}
