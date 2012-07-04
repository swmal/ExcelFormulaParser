using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class Ranges
    {
        private readonly Dictionary<Guid, List<RangeAddress>> _rangeAddresses = new Dictionary<Guid, List<RangeAddress>>();

        public virtual void Add(ParsingScope parsingScope, string rangeAddress)
        {
            Add(parsingScope, RangeAddress.Parse(rangeAddress));
        }

        public virtual void Add(ParsingScope parsingScope, RangeAddress rangeAddress)
        {
            if (!_rangeAddresses.ContainsKey(parsingScope.ScopeId))
            {
                _rangeAddresses.Add(parsingScope.ScopeId, new List<RangeAddress>());
            }
            _rangeAddresses[parsingScope.ScopeId].Add(rangeAddress);
        }

        public virtual void Clear()
        {
            _rangeAddresses.Clear();
        }

        public virtual void CheckCircularReference(ParsingScope scope, RangeAddress rangeAddress)
        {
            var conflictingScope = (from r in _rangeAddresses
                                 where r.Key != scope.ScopeId
                                 && r.Value.Exists(x => x.CollidesWith(rangeAddress))
                                 select r).FirstOrDefault();
            if(conflictingScope.Value != null)
            {
                var conflictingRange = conflictingScope.Value.Find(x => x.CollidesWith(rangeAddress));
                var errMsg = string.Format("Circular reference detected between {0} and {1}", conflictingRange, rangeAddress);
                throw new CircularReferenceException(errMsg);
            }
        }
    }
}
