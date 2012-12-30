using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalcCell
    {
        private readonly Dictionary<object, CalcRelation> _relationsTo = new Dictionary<object, CalcRelation>();
        private readonly Dictionary<object, CalcRelation> _relationsFrom = new Dictionary<object, CalcRelation>();
        private readonly CalcChainContext _context;

        private CalcCell(string address, CalcChainContext context)
        {
            Require.That(context).Named("context").IsNotNull();
            Id = context.IdProvider.NewId();
            _context = context;
            Address = address;
        }

        public static CalcCell Create(string address, CalcChainContext context)
        {
            return new CalcCell(address, context);
        }

        public IEnumerable<CalcRelation> RelationsTo
        {
            get { return _relationsTo.Values; }
        }

        public IEnumerable<CalcRelation> RelationsFrom
        {
            get { return _relationsFrom.Values; }
        }

        public object Id { get; private set; }

        public string Address { get; private set; }

        public CalcRelation AddRelationTo(CalcCell other, CalcChain chain)
        {
            if (_relationsTo.ContainsKey(other.Id)) return _relationsTo[other.Id];
            var rel = new CalcRelation { CalcCellId = other.Id, CalcChainId = chain.Id };
            _relationsTo.Add(other.Id, rel);
            return rel;
        }

        public CalcRelation AddRelationFrom(CalcCell other, CalcChain chain)
        {
            if (_relationsFrom.ContainsKey(other.Id)) return _relationsFrom[other.Id];
            var rel = new CalcRelation { CalcCellId = other.Id, CalcChainId = chain.Id };
            _relationsFrom.Add(other.Id, rel);
            return rel;
        }

        public object GetCalcChainId()
        {
            CalcChain resultChain = null;
            foreach (var rel in _relationsFrom)
            {
                var chain = _context.GetCalcChain(rel.Value.CalcChainId);
                if (resultChain == null || chain.Count > resultChain.Count)
                {
                    resultChain = chain;
                }
            }
            return resultChain != null ? resultChain.Id : null;
        }

        public override string ToString()
        {
            return Address ?? "??";
        }
    }
}
